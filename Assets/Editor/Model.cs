using System;
using System.Collections.Generic;

[Serializable]
public class Root
{
    public List<Language> LanguageList;
    public Root()
    {
        LanguageList = new List<Language>();
    }
}
[Serializable]
public class Language
{
    public string Key;
    public string ChineseSimplified;
    public string ChineseTraditional;
    public string English;
    public Language(string key, string chineseSimplified, string chineseTraditional, string english)
    {
        Key = key;
        ChineseSimplified = chineseSimplified;
        ChineseTraditional = chineseTraditional;
        English = english;
    }
}