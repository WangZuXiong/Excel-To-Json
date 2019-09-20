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
        //Row[][] Row[行][列]
        Root root = new Root();
        for (int i = 1; i < dataTable.Rows.Count; i++)
        {
            var array = dataTable.Rows[i].ItemArray;
            var temp = new Language(array[0].ToString(), array[1].ToString(), array[2].ToString(), array[3].ToString());
            root.LanguageList.Add(temp);
        }
        //生成Json字符串
        var jsonStr = JsonUtility.ToJson(root, true);
        //写入文件
        WriteFile(jsonStr, _outputPath);
        AssetDatabase.Refresh();
        Debug.LogError(">>>转换成功");
    }

    private void WriteFile(string str,string path)
    {
        using (FileStream fileStream = new FileStream(path, FileMode.Create, FileAccess.Write))
        {
            using (TextWriter textWriter = new StreamWriter(fileStream, Encoding.UTF8))
            {
                textWriter.Write(str);
            }
        }
    }
}