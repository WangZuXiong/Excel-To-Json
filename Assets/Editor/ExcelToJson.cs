using Excel;
using System;
using System.Data;
using System.IO;
using System.Text;
using UnityEditor;
using UnityEngine;

public class SimpleTool : EditorWindow
{
    private static string _inputPath;
    private static string _outputPath;

    private void OnGUI()
    {
        GUILayout.Label("Input Path:");
        _inputPath = GUILayout.TextField(_inputPath);
        GUILayout.Label("Output Path:");
        _outputPath = GUILayout.TextField(_outputPath);
        if (GUILayout.Button("转换"))
        {
            Change();
            //将path保存到本地
            SavePath();
        }
    }

    [MenuItem("Tool/excel to json")]
    public static void ShowWindow()
    {
        //加载path
        LoadPath();
        var window = GetWindow(typeof(SimpleTool), false);
        window.titleContent = new GUIContent("excel转为json文件", "Set Path");
        window.Show();
    }

    private static void LoadPath()
    {
        _inputPath = PlayerPrefs.GetString(Application.productName + "INPUT_PATH", string.Empty);
        _outputPath = PlayerPrefs.GetString(Application.productName + "OUTPUT_PATH", string.Empty);
    }

    private void SavePath()
    {
        PlayerPrefs.SetString(Application.productName + "INPUT_PATH", _inputPath);
        PlayerPrefs.SetString(Application.productName + "OUTPUT_PATH", _outputPath);
    }

    private void Change()
    {
        if (!File.Exists(_inputPath))
        {
            throw new Exception("不存在文件：" + _inputPath);
        }

        FileStream mStream;
        try
        {
            //文件流
            mStream = File.Open(_inputPath, FileMode.Open, FileAccess.Read);
        }
        catch
        {
            throw new Exception("文件正在被占用中:" + _inputPath);
        }


        DataTable dataTable;
        //列数
        var index = 0;
        //转化为excel
        IExcelDataReader factoy = ExcelReaderFactory.CreateOpenXmlReader(mStream);
        //表格数据集合
        var dataSet = factoy.AsDataSet();

        if (dataSet.Tables.Count < 1)
        {
            return;
        }
        //获取第一个数据表
        dataTable = dataSet.Tables[0];
        for (int i = 0; i < dataTable.Columns.Count; i++)
        {
            if (dataTable.Rows[0][i].ToString().Equals("Key"))
                index = i;
        }
        //Row[][] Row[行][列]
        var count = dataTable.Rows.Count;

        //Debug.LogError(dataTable.Rows[0]);
        Root root = new Root();
        for (int i = 1; i < count; i++)
        {
            var key = dataTable.Rows[i][index].ToString();
            var chineseSimplified = dataTable.Rows[i][index + 1].ToString();
            var chineseTraditional = dataTable.Rows[i][index + 2].ToString();
            var english = dataTable.Rows[i][index + 3].ToString();
            var temp = new Language(key, chineseSimplified, chineseTraditional, english);
            root.LanguageList.Add(temp);
        }

        //生成Json字符串
        //var json = JsonConvert.SerializeObject(root, Formatting.Indented);

        var json = JsonUtility.ToJson(root, true);

        //写入文件
        using (FileStream fileStream = new FileStream(_outputPath, FileMode.Create, FileAccess.Write))
        {
            using (TextWriter textWriter = new StreamWriter(fileStream, Encoding.UTF8))
            {
                textWriter.Write(json);
            }
        }
        AssetDatabase.Refresh();
        Debug.LogError(">>>转换成功");
    }
}