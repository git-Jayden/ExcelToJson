#if UNITY_EDITOR 
using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Reflection;
using System.Text;
using UnityEditor;
using UnityEngine;

public class ExcelToJson : Editor
{
    /// <summary>
    /// Excel转Jsoun并保存Json
    /// </summary>
    [MenuItem("Assets/ExcelToJson")]
    static void DoXlsxToJson()
    {
        List<string> names = new List<string>();
        var o = Selection.GetFiltered(typeof(UnityEngine.Object), SelectionMode.DeepAssets);
        foreach (UnityEngine.Object obj in o)
        {
            string filePath = AssetDatabase.GetAssetPath(obj);

            if (filePath.IndexOf('.') > 0)
            {
                string[] s = filePath.Split('.');
                if (s[1] == "xlsx")
                    names.Add(filePath);
            }
        }
        if (names.Count == 0)
            return;
        foreach (var item in names)
        {
            ReadSingleExcel(item);
           
        }
    }
    /// <summary>
    /// 读取单张Excel
    /// </summary>
    /// <param name="xlsxPath"></param>
    /// <param name="jsonPath"></param>
    public static void ReadSingleExcel(string xlsxPath, string jsonPath = null)
    {
        string dataPath = UnityEngine.Application.dataPath;

        if (jsonPath == null)
        {
            jsonPath = dataPath + "/Resources/Json";
            if (!Directory.Exists(jsonPath))
                Directory.CreateDirectory(jsonPath);
        }
        string xlsxName = FileTool.GetFileName(xlsxPath);
        jsonPath += "/" + xlsxName;

        if (!File.Exists(xlsxPath))
        {
            UnityEngine.Debug.LogError("不存在该Excel！");
            return;
        }
    
        using (Stream stream = File.Open(xlsxPath, FileMode.Open, FileAccess.Read))
        {
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            DataSet result = excelReader.AsDataSet();


            // 读取第一张工资表
            DataTableCollection tc = result.Tables;
            for (int i = 0; i < tc.Count; i++)
            {
                ReadSingleSheet(result.Tables[i], jsonPath);
            }
            Debug.Log(xlsxName+ ".xlsx" + "转换Json成功！");
        }
    }
    /// <summary>
    /// 读取一个工作表的数据
    /// </summary>
    /// <param name="type">要转换的struct或class类型</param>
    /// <param name="dataTable">读取的工作表数据</param>
    /// <param name="jsonPath">存储路径</param>
    private static void ReadSingleSheet(System.Data.DataTable dataTable, string jsonPath)
    {

        Assembly assembly = Assembly.LoadFile(Environment.CurrentDirectory + @"\Library\ScriptAssemblies\Assembly-CSharp.dll"); // 获取当前程序集 
        Type type = assembly.GetType(dataTable.TableName); // 创建类的实例，返回为 object 类型，需要强制类型转换
        //string TableNameJson = jsonPath + "/" + TableName + ".json";
        //Type type = typeof(T);
        int rows = dataTable.Rows.Count;
        int Columns = dataTable.Columns.Count;
        // 工作表的行数据
        DataRowCollection collect = dataTable.Rows;
        // xlsx对应的数据字段，规定是第二行
        string[] jsonFileds = new string[Columns];
        // 要保存成Json的obj
        List<object> objsToSave = new List<object>();
        for (int i = 0; i < Columns; i++)
        {
            jsonFileds[i] = collect[1][i].ToString();

        }
        // 从第三行开始
        for (int i = 3; i < rows; i++)
        {
            // 生成一个实例
            object objIns = Activator.CreateInstance(type);

            for (int j = 0; j < Columns; j++)
            {
                // 获取字段
                FieldInfo field = type.GetField(jsonFileds[j]);
                if (field != null)
                {
                    object value = null;
                    try // 赋值
                    {
                        value = Convert.ChangeType(collect[i][j], field.FieldType);
                    }
                    catch (InvalidCastException e) // 一般到这都是Int数组，当然还可以更细致的处理不同类型的数组
                    {
                        Console.WriteLine(e.Message);
                        string str = collect[i][j].ToString();
                        string[] strs = str.Split(',');
                        int[] ints = new int[strs.Length];
                        for (int k = 0; k < strs.Length; k++)
                        {
                            ints[k] = int.Parse(strs[k]);
                        }
                        value = ints;
                    }
                    field.SetValue(objIns, value);
                }
                else
                {
                    UnityEngine.Debug.LogFormat("有无法识别的字符串：{0}", jsonFileds[j]);
                }
            }
            objsToSave.Add(objIns);
        }
        string content =Newtonsoft.Json.JsonConvert.SerializeObject(objsToSave);
        jsonPath += "_"+ dataTable.TableName + ".json";
        FileTool.SaveFile(content, jsonPath);    
    }

}
#endif