using UnityEngine;
using UnityEditor;
using System.IO;
using System.Collections.Generic;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System;
using ExcelConverter;
using SFB;
/// <summary>
//NOTE - 만든이 : 백수진모씨
//NOTE - 소속 : 15pluto
//NOTE - 만든날짜 : 2022-11.12
//NOTE - 문의 : yakmo94@naver.com
/// </summary>
[CustomEditor(typeof(ExcelLoader))]
public class CExcelLoaderInspector : Editor
{
    //! 클레스명을 key값으로 하고 읽을 파일명을 value값으로 하는 딕셔너리 .
    static Dictionary<string, string> m_CsDic = new Dictionary<string, string>();
    // //! 클레스 명을 key값으로 하고 클레스를 구성하는 변수 명을 value값으로 하는 딕셔너리. 
    // static Dictionary<string, List<string>> ExcelNameDic = new Dictionary<string, List<string>>();
    //! 엑셀에서 sheet이름을 읽어서 클레스명과 데이터 파일명으로 사용하는데 같은 이름이 있는지 확인을 하기 위해서 sheet이름을 key로 하고 같은 이름의 파일 갯수를 value로 하는 익셔너리.
    static Dictionary<string, int> m_SaveFileNameDic = new Dictionary<string, int>();
    /// <summary>
    /// This function is called when the object becomes enabled and active.
    /// 이때 저장한 데이터를 불러오자.
    /// </summary>
    private void OnEnable()
    {
        string saveFileName = Path.Combine(Application.persistentDataPath,"excelLoader.txt") ;
        if(System.IO.File.Exists(saveFileName))
        {
            ExcelLoader.I.m_Paths.Clear() ;
            System.IO.StreamReader fp = new StreamReader(saveFileName) ;
            //몇개의 엑셀파일이 저장되었는지 확인한다.
            int count = Convert.ToInt32(fp.ReadLine());
            for(int i = 0 ; i < count ; i ++)
            {
                //이름을 읽어 온다.
                var path = fp.ReadLine(); 
                //중간에 삭제 했는지 확인 해보자
                if(!File.Exists(path)) continue ;
                ExcelLoader.I.m_Paths.Add(path) ;
            }
            fp.Close() ;
        }
    }
    /// <summary>
    //REVIEW 외부에서 읽어온 엑셀 파일은 프로젝트를 껐다가 키면 사라진다 그래서 그걸 따로 저장했다가 다시 불러주자.
    /// </summary>
    void SaveFileNames()
    {
        string saveFileName = Path.Combine(Application.persistentDataPath,"excelLoader.txt") ;
        var fp = new StreamWriter(saveFileName);
        var paths = ExcelLoader.I.m_Paths ;
        //총 갯수를 저장 한다.
        fp.WriteLine(paths.Count.ToString()) ;
        foreach(var item in paths)
        {
            //파일 이름을 저장 해준다.
            fp.WriteLine(item) ;
        }
        fp.Close();
    }
    public override void OnInspectorGUI()
    {
        base.OnInspectorGUI();
        GUILayout.Space(10);
        if(ExcelLoader.I.m_Paths.Count > 0)
        {
            GUILayout.Label("★외부에서 추가된 파일들★");
        }
        GUIStyle myStyle = new GUIStyle(EditorStyles.textField);
        myStyle.fontSize = 12;
        myStyle.wordWrap = true;
        foreach (var item in ExcelLoader.I.m_Paths)
        {
            //  ExcelLoader.I.m_Paths.Clear() ;
            GUI.backgroundColor = Color.black;
            GUILayout.BeginHorizontal("box");
            GUI.backgroundColor = Color.white;
            GUILayout.TextArea(item, myStyle);
            GUI.backgroundColor = Color.red;
            //remove
            if (GUILayout.Button("삭제", GUILayout.Width(40), GUILayout.MinWidth(40)))
            {
                ExcelLoader.I.m_Paths.Remove(item);
                SaveFileNames() ;
                break;
            }
            
            GUI.backgroundColor = Color.white;
            GUILayout.EndHorizontal();
        }
        GUILayout.Space(10) ;
        GUI.backgroundColor = Color.black;
        GUILayout.BeginHorizontal("box");
        GUI.backgroundColor = Color.white;
        //FILE OPEN
        if (GUILayout.Button("외부 파일 선택", GUILayout.Height(30)))
        {
            var extensions = new[] {new ExtensionFilter("MS Excel", "xlsx", "xls" )};
            var paths = StandaloneFileBrowser.OpenFilePanel("MS Excel 파일 열기", GetPathForOpenPanel(), extensions, true);
            foreach (var item in paths)
            {
                if (!item.ToLower().EndsWith(".xlsx") && !item.ToLower().EndsWith(".xls")) continue;
                if (!System.IO.File.Exists(item)) continue;
                if (ExcelLoader.I.m_Paths.Contains(item)) continue;
                ExcelLoader.I.m_Paths.Add(item);
                Debug.Log(item);
            }
            SaveFileNames() ;
        }
        GUI.backgroundColor = Color.magenta;
        //FOLDER OPEN
        if (GUILayout.Button("외부 폴더 선택", GUILayout.Height(30)))
        {
            var folder = EditorUtility.OpenFolderPanel("MS Excel 파일 열기", GetPathForOpenPanel(), "");
            ///폴더 선택창을 눌렀을때 취소 버튼을 누르면 이름이 빈값으로 들어 온다.제대로 선택을 했는지 확인을 먼저 해야 에러가 발생하지 않는다.
            if (Directory.Exists(folder))
            {
                System.IO.DirectoryInfo di = new DirectoryInfo(folder);
                foreach (var item in di.GetFiles())
                {
                    //'~'가 포함된건 엑셀 파일이 열렸을때 생성되는 임시 파일이니까 읽지 말자.
                    if (item.Name.Contains("~")) continue;
                    if (!item.Name.ToLower().EndsWith(".xlsx") && !item.Name.ToLower().EndsWith(".xls")) continue;
                    if (ExcelLoader.I.m_Paths.Contains(item.FullName)) continue;
                    ExcelLoader.I.m_Paths.Add(item.FullName);
                    Debug.Log(item.FullName);
                }
                SaveFileNames();
            }
        }
        GUILayout.EndHorizontal();
        string saveFileName = Path.Combine(Application.persistentDataPath,"excelLoader.txt") ;
        GUI.backgroundColor = Color.red;
        if (GUILayout.Button("세이브파일 위치 복사", GUILayout.Height(20)))
        {
            GUIUtility.systemCopyBuffer =  Application.persistentDataPath ;//Path.Combine(Application.persistentDataPath,"excelLoader.txt") ;
            Debug.Log($"{Application.persistentDataPath} 클립보드에 복사가 완료 되었습니다.");
        }
        if (ExcelLoader.DOWNLOADPATH.Length == 0) return;
        if (ExcelLoader.CODEPATH.Length == 0) return;
        if (ExcelLoader.ExcelCount == 0 && ExcelLoader.I.m_Paths.Count == 0) return;        
        GUI.backgroundColor = Color.green;
        GUILayout.Space(2);
        //START DATA CREATE
        if (GUILayout.Button("데이터 및 코드 생성", GUILayout.Height(40)))
        {
            StartDataCreate();
        }
    }
    //REVIEW openfile이나 oepnfolder를 할때 처음 보여질 폴더위치를 알기 위해서 외부에서 불러온 엑셀 파일의 맨 마지막 녀석의 폴더명을 가져와서 리턴 해준다.
    string GetPathForOpenPanel()
    {
        var path = "";
        if (ExcelLoader.I.m_Paths.Count > 0)
        {
            path = ExcelLoader.I.m_Paths[ExcelLoader.I.m_Paths.Count - 1];
            FileInfo fi = new FileInfo(ExcelLoader.I.m_Paths[ExcelLoader.I.m_Paths.Count - 1]);
            var idx = path.LastIndexOf(fi.Name);
            path = path.Remove(idx, path.Length - idx);
        }
        return path;
    }
    /// <summary>
    //REVIEW 데이터 생성을 하기 위해서 호출 되는 함수
    /// </summary>
    void StartDataCreate()
    {
        //나중에 코드를 만들때 메니저에 필요한 변수를 만들때 사용되는 딕셔너리를 초기화 해준다.
        m_CsDic = new Dictionary<string, string>();
        //시트이름이 겹칠때를 대비하여 뒤에 숫자를 붙이기 위해 필요한 변수
        m_SaveFileNameDic = new Dictionary<string, int>();
        var loader = target as ExcelLoader;
        CreateBaseCsFile(ExcelLoader.CODEPATH);
        // 엑셀 데이터들을 읽어 와서 하나씩 읽어서 데이터 파일과 코드를 만들어 주자.
        foreach (var item in loader.ExcelDatas)
        {
            try
            {
                var fullName = AssetDatabase.GetAssetPath(item);
                string excelName = Path.GetFileNameWithoutExtension(fullName);
                CreateExcelDataOne(fullName);
            }
            catch { }
        }
        foreach (var item in ExcelLoader.I.m_Paths)
        {
            CreateExcelDataOne(item);
        }
        CreateManagerFile(ExcelLoader.CODEPATH, m_CsDic, ExcelLoader.RESOURCEPATH);
        AssetDatabase.SaveAssets();
        AssetDatabase.Refresh();
        
    }
    /// <summary>
    //REVIEW  Resources.Load를 이용하여 데이터를 읽기 위해서는 Resources의 하위 폴더의 경로만이 필요하다 그걸 찾기 위해서 만든 함수
    /// </summary>
    /// <param name="respath">리소스를 저장할 위치</param>
    /// <returns></returns>
    string GetPathExceptResoureces(string respath)
    {
        int idx = respath.LastIndexOf("Resources") ;
        respath = respath.Remove(0,idx + "Resources".Length) ;
        if(respath.Length > 0 && respath[0] == '/')
        {
                respath = respath.Remove(0,1) ;
        } 
        return respath ;
    }
    /// <summary>
    //REVIEW 리소스를 읽고 사용하는걸 관리하는 클레스를 만든다. 무식하게 만든다.
    /// </summary>
    /// <param name="filePath">메니저 파일을 저장할 위치</param>
    /// <param name="NameDic">클래스명을 Key로 사용하고 파일명을 Value로 사용하는 딕셔너리</param>
    /// <param name="resPath">리소스 저장위치</param>
    void CreateManagerFile(string filePath, Dictionary<string, string> NameDic, string resPath)
    {
        string managerName = ExcelLoader.I.Manager클레스이름;
        var path = Path.Combine(filePath, $"{managerName}.cs");
        if (System.IO.File.Exists(path))
        {
            System.IO.File.Delete(path);
        }
        System.IO.StreamWriter fp = new System.IO.StreamWriter(path);
        fp.WriteLine("using UnityEngine;");
        fp.WriteLine("using System.Collections.Generic;");
        fp.WriteLine("using System;");
        fp.WriteLine("using System.IO;");
        fp.WriteLine();
        fp.WriteLine("namespace MsExcelLoader");
        fp.WriteLine("{");//namespace
        CreateAddTab(fp, 1); fp.WriteLine("[System.Serializable]");
        CreateAddTab(fp, 1); fp.WriteLine($"public class {managerName} : CSingleton<{managerName}>");
        CreateAddTab(fp, 1); fp.WriteLine("{");//start class
        CreateAddTab(fp, 2); fp.WriteLine($"public string CODEPATH = \"{GetPathExceptResoureces(ExcelLoader.RESOURCEPATH)}\";");
        //변수 만들기
        CreateAddTab(fp, 2); fp.WriteLine($"Dictionary<string,List<CCDNBase>>  m_ItemList = new Dictionary<string, List<CCDNBase>>() ;");
        CreateAddTab(fp, 2); fp.WriteLine($"Dictionary<string,Dictionary<string,CCDNBase>>  m_ItemDic = new Dictionary<string, Dictionary<string, CCDNBase>>() ;");
        CreateAddTab(fp, 2); fp.WriteLine("#region Property and Getter");
        CreateAddTab(fp, 2); fp.WriteLine($"public CCDNBase this[string filename, int idx]");
        CreateAddTab(fp, 2); fp.WriteLine("{");
        CreateAddTab(fp, 3); fp.WriteLine($"get=>m_ItemList[filename][idx] ;");
        CreateAddTab(fp, 2); fp.WriteLine("}");
        CreateAddTab(fp, 2); fp.WriteLine($"public CCDNBase this[string filename, string key]");
        CreateAddTab(fp, 2); fp.WriteLine("{");
        CreateAddTab(fp, 3); fp.WriteLine($"get=>m_ItemDic[filename][key] ;");
        CreateAddTab(fp, 2); fp.WriteLine("}");
        CreateAddTab(fp, 2); fp.WriteLine($"public List<CCDNBase> getDataList(string filename)");
        CreateAddTab(fp, 2); fp.WriteLine("{");
        CreateAddTab(fp, 3); fp.WriteLine($"return m_ItemList[filename] ;");
        CreateAddTab(fp, 2); fp.WriteLine("}");
        CreateAddTab(fp, 2); fp.WriteLine("public Dictionary<string,CCDNBase> GetDataDic(string filename)");
        CreateAddTab(fp, 2); fp.WriteLine("{");
        CreateAddTab(fp, 3); fp.WriteLine("return m_ItemDic[filename] ;");
        CreateAddTab(fp, 2); fp.WriteLine("}");
        CreateAddTab(fp, 2); fp.WriteLine("#endregion");
        fp.WriteLine();
        CreateAddTab(fp, 2); fp.WriteLine("public override void Init()");
        CreateAddTab(fp, 2); fp.WriteLine("{");//function
        foreach (var data in NameDic)
        {
            var className = data.Key;//"C" + char.ToUpper(data.m_Name[0]) + data.m_Name.Substring(1).ToLower();
            CreateAddTab(fp, 3); fp.WriteLine($"MakeData<{className}>(\"{data.Value}\");");
        }
        // CreateAddTab(fp, 3); fp.WriteLine("CExcelManager.I.m_EndCallback?.Invoke() ;");
        CreateAddTab(fp, 2); fp.WriteLine("}");//function
        CreateAddTab(fp, 2); fp.WriteLine("public void MakeData<T>(string filename) where T :  CCDNBase ,new ()");
        CreateAddTab(fp, 2); fp.WriteLine("{");//function
        CreateAddTab(fp, 3); fp.WriteLine("m_ItemList[filename] = new List<CCDNBase>() ;");
        CreateAddTab(fp, 3); fp.WriteLine("m_ItemDic[filename] = new Dictionary<string,CCDNBase>() ;");
        //여기문제다
        //Path.Combine(CODEPATH, filename)
        CreateAddTab(fp, 3); fp.WriteLine("TextAsset asset = Resources.Load<TextAsset>(Path.Combine(CODEPATH, filename));");
        CreateAddTab(fp, 3); fp.WriteLine("var  fp = new System.IO.BinaryReader(new MemoryStream(asset.bytes));");
        //CreateAddTab(fp, 3); fp.WriteLine("var fp = CBinaryReader.Create(Path.Combine(CODEPATH, filename), false);"); //파일을 읽어 올때 문제가 된다. 여기를 어떻게 해야 한다.
        //CreateAddTab(fp, 3); fp.WriteLine("var fp = CExcelManager.I.GetFilePoint(filename);"); //파일을 읽어 올때 문제가 된다. 여기를 어떻게 해야 한다.
        CreateAddTab(fp, 3); fp.WriteLine("var CountForLoad = fp.ReadInt32() ;");
        CreateAddTab(fp, 3); fp.WriteLine("for(int i = 0 ; i < CountForLoad ; i ++)");
        CreateAddTab(fp, 3); fp.WriteLine("{");//for
        CreateAddTab(fp, 4); fp.WriteLine("var data = new T() ;");
        CreateAddTab(fp, 4); fp.WriteLine("data.Init(fp) ;");
        CreateAddTab(fp, 4); fp.WriteLine("m_ItemList[filename].Add (data);");
        CreateAddTab(fp, 4); fp.WriteLine("m_ItemDic[filename][data.KEY] = data;");
        CreateAddTab(fp, 3); fp.WriteLine("}");//endfor
        CreateAddTab(fp, 3); fp.WriteLine("fp.Close() ;"); //파일을 읽어 올때 문제가 된다. 여기를 어떻게 해야 한다.
        CreateAddTab(fp, 2); fp.WriteLine("}");//endfunction
        CreateAddTab(fp, 1); fp.WriteLine("}");//start class
        fp.WriteLine("}");//namespace
        fp.Close();
    }
    /// <summary>
    //REVIEW 엑셀에 시트이름을 파일이름과 cs파일 이름으로 사용을 하는데 같은 이름이 있을 경우 뒤에 숫자를 붙여서 진짜 이름을 만들어준다. 그때 사용되는 함수
    /// </summary>
    /// <param name="name">엑셀에서 읽어온 시트 이름</param>
    /// <returns></returns>
    static string GetRealFileName(string name)
    {
        if (m_SaveFileNameDic.ContainsKey(name))
        {
            m_SaveFileNameDic[name]++;
            return $"{name}{m_SaveFileNameDic[name]}";
        }
        m_SaveFileNameDic[name] = 0;
        return name;
    }
    static void _Write(System.IO.BinaryWriter fp, CType type, string data, Dictionary<string, List<string>> enumDic)
    {
        var typeName = type.m_Name.Replace(" ", "").ToLower();
        if (typeName == "enum")
        {
            var enumList = enumDic[type.m_Data];
            int idx = enumList.IndexOf(data);
            fp.Write(idx);
        }
        else if (typeName.ToLower() == "list")
        {
            var listType = type.m_Data;
            var datas = data.Split(',');
            fp.Write((int)datas.Length);
            foreach (var item in datas)
            {
                WriteOneData(fp, listType, item);
            }
        }
        else
        {
            WriteOneData(fp, type.m_Name, data);
        }
    }
    /// <summary>
    //REVIEW 데이터를 저장할때 사용되는 함수
    /// </summary>
    /// <param name="fp">파일포인터</param>
    /// <param name="type">데이터 타입</param>
    /// <param name="data">데이터</param>
    static void WriteOneData(System.IO.BinaryWriter fp, string type, string data)
    {
        try
        {
            switch (type.ToLower())
            {
                case "bool":
                    fp.Write(bool.Parse(data));
                    break;
                case "byte":
                    fp.Write(byte.Parse(data));
                    break;
                case "char":
                    fp.Write(Convert.ToInt16(data));
                    break;
                case "short":
                    fp.Write(short.Parse(data));
                    break;
                case "uint16":
                    fp.Write(Int16.Parse(data));
                    break;
                case "int":
                    fp.Write(int.Parse(data));
                    break;
                case "uint32":
                    fp.Write(UInt32.Parse(data));
                    break;
                case "int64":
                    fp.Write(Int64.Parse(data));
                    break;
                case "uint64":
                    fp.Write(UInt64.Parse(data));
                    break;
                case "float":
                    fp.Write(float.Parse(data));
                    break;
                case "double":
                    fp.Write(double.Parse(data));
                    break;
                default://!default는 스트링이다 .
                    fp.Write(data);
                    break;
            }
        }
        catch
        {
            Debug.LogError(type + " " + data);
        }
    }
    /// <summary>
    //REVIEW 코드를 만들때 탭을 추가 해서 코드를 이쁘게 만들때 사용한다.
    /// </summary>
    /// <param name="fp">System.IO.StreamWriter</param>
    /// <param name="count">저장할 탭 갯수</param>
    static void CreateAddTab(System.IO.StreamWriter fp, int count)
    {
        for (int i = 0; i < count; i++)
        {
            fp.Write("\t");
        }
    }
    /// <summary>
    //REVIEW 코드를 만들때 데이터 타입에 따라서 읽는 부분을 다르게 해줘야 하는데 그 부분을 만들어 주는 함수
    /// </summary>
    /// <param name="type">데이터 타입</param>
    /// <param name="enumType">enum타입</param>
    /// <returns></returns>
    static string GetReadDataByType(string type, string enumType)
    {
        type = type.ToLower().Replace(" ", "");
        if (type == "char")
        {
            return "fp.ReadChar();";
        }
        if (type == "short")
        {
            return "fp.ReadInt16();";

        }
        if (type == "int")
        {
            return "fp.ReadInt32();";
        }
        if (type == "int64")
        {
            return "fp.ReadInt64();";
        }//-----------
        if (type == "byte")
        {
            return "fp.ReadByte();";
        }
        if (type == "uint16")
        {

            return "fp.ReadUInt16();";
        }
        if (type == "uint32")
        {
            return "fp.ReadUInt32();";
        }
        if (type == "uint64")
        {
            return "fp.ReadUInt64();";
        }

        if (type == "float")
        {
            return "fp.ReadSingle();";
        }
        if (type == "double")
        {
            return "fp.ReadDouble();";
        }
        if (type == "enum")
        {
            return $"({enumType})fp.ReadInt32();";
        }
        if (type == "bool")
        {
            return "fp.ReadByte() == 1 ? true : false;";
        }
        return "fp.ReadString() ;";
    }
    static void CreateBaseCsFile(string codepath)
    {
        var path = Path.Combine(codepath, "CCDNBase.cs");
        if (System.IO.File.Exists(path))
        {
            System.IO.File.Delete(path);
        }
        System.IO.StreamWriter fp = new System.IO.StreamWriter(path);

        fp.WriteLine("using UnityEngine;");
        fp.WriteLine("using System.Collections.Generic;");
        fp.WriteLine("using System;");
        fp.WriteLine("using System.IO;");
        fp.WriteLine();
        fp.WriteLine("namespace MsExcelLoader");
        fp.WriteLine("{");//namespace
        CreateAddTab(fp, 1); fp.WriteLine($"public class CCDNBase");
        CreateAddTab(fp, 1); fp.WriteLine("{");//start class
        CreateAddTab(fp, 2); fp.WriteLine("public virtual string KEY");
        CreateAddTab(fp, 2); fp.WriteLine("{");
        CreateAddTab(fp, 3); fp.WriteLine("get =>\"CCDNBase\";");
        CreateAddTab(fp, 2); fp.WriteLine("}");
        CreateAddTab(fp, 2); fp.WriteLine("public virtual string Name");
        CreateAddTab(fp, 2); fp.WriteLine("{");
        CreateAddTab(fp, 3); fp.WriteLine("get =>\"CCDNBase\";");
        CreateAddTab(fp, 2); fp.WriteLine("}");
        CreateAddTab(fp, 2); fp.WriteLine("public virtual void Init(System.IO.BinaryReader fp){}");
        CreateAddTab(fp, 1); fp.WriteLine("}");//start class
        fp.WriteLine("}");//namespace
        fp.Close();
    }
    /// <summary>
    //REVIEW cs파일을 만든다.
    /// </summary>
    /// <param name="data">cs파일을 만드는데 필요한 데이터</param>
    static void CreateCode(CExcelData data)
    {
        var path = data.m_CSPath;
        if (System.IO.File.Exists(path))
        {
            System.IO.File.Delete(path);
        }
        System.IO.StreamWriter fp = new System.IO.StreamWriter(path);
        fp.WriteLine("using UnityEngine;");
        fp.WriteLine("using System.Collections.Generic;");
        fp.WriteLine("using System;");
        fp.WriteLine();
        fp.WriteLine("namespace MsExcelLoader");
        var className = data.m_Name;//"C" + char.ToUpper(data.m_Name[0]) + data.m_Name.Substring(1).ToLower();
        fp.WriteLine("{");//namespace
        CreateAddTab(fp, 1); fp.WriteLine("[System.Serializable]");
        //NOTE -  CCDNBase
        CreateAddTab(fp, 1); fp.WriteLine($"public class {className} : CCDNBase");
        CreateAddTab(fp, 1); fp.WriteLine("{");//start class
                                               //enum 만들기
        foreach (var item in data.m_EnumDic)
        {
            CreateAddTab(fp, 2); fp.Write($"public enum {item.Key} {{ NONE = -1");//start class    
            for (int i = 0; i < item.Value.Count; i++)
            {
                //    if(i != 0)
                {
                    fp.Write(",");
                }
                fp.Write(item.Value[i]);
            }
            fp.Write(",MAX}");
            fp.WriteLine();
        }
        #region 프로퍼티
        CreateAddTab(fp, 2); fp.WriteLine("public override string Name");
        CreateAddTab(fp, 2); fp.WriteLine("{");
        CreateAddTab(fp, 3); fp.WriteLine($"get =>\"{className}\";");
        CreateAddTab(fp, 2); fp.WriteLine("}");
        CreateAddTab(fp, 2); fp.WriteLine($"public override string KEY");
        CreateAddTab(fp, 2); fp.WriteLine("{");
        CreateAddTab(fp, 3); fp.WriteLine($"get=> {data.m_ValueNames[0]}.ToString();");
        CreateAddTab(fp, 2); fp.WriteLine("}");

        #endregion
        for (int i = 0; i < data.m_ValueNames.Count; i++)
        {
            var type = data.m_DataTypes[i].m_Name;
            var name = data.m_ValueNames[i];
            CreateAddTab(fp, 2);
            if (type == "enum")
            {
                fp.WriteLine($"public {data.m_DataTypes[i].m_Data} {name}{{get; private set; }}");
            }
            else if (type.ToLower() == "list")
            {
                fp.WriteLine($"public List<{data.m_DataTypes[i].m_Data}> {name} {{get; private set; }} = new List<{data.m_DataTypes[i].m_Data}>();");
            }
            else
            {
                fp.WriteLine($"public {type} {name} {{get; private set; }}");
            }
        }
        //--------------------------------------------------------------------------------------------------------------------
        CreateAddTab(fp, 2); fp.WriteLine($"public  override void Init(System.IO.BinaryReader fp)");
        CreateAddTab(fp, 2); fp.WriteLine("{");//end class
        bool needcountformakefile = false;
        for (int i = 0; i < data.m_DataTypes.Count; i++)
        {
            var type = data.m_DataTypes[i];
            if (type.m_Name.ToLower() == "list")
            {
                needcountformakefile = true;
                break;
            }
        }
        if (needcountformakefile)
        {
            CreateAddTab(fp, 3); fp.WriteLine("int countForMakeFile = 0;");//갯수를 읽어와서
        }
        for (int i = 0; i < data.m_DataTypes.Count; i++)
        {
            var datatype = data.m_DataTypes[i];
            var TitleName = data.m_ValueNames[i];
            if (datatype.m_Name.ToLower() == "list")
            {
                var Type = datatype.m_Data; CreateAddTab(fp, 3); fp.WriteLine("countForMakeFile = fp.ReadInt32();");//갯수를 읽어와서
                CreateAddTab(fp, 3); fp.WriteLine("for( int i = 0 ; i < countForMakeFile ; i ++)");
                CreateAddTab(fp, 3); fp.WriteLine("{");
                var value = GetReadDataByType(Type, "");//value
                CreateAddTab(fp, 4); fp.WriteLine($"var value = {value}");
                var str = $"{TitleName}.Add(value);";
                CreateAddTab(fp, 4); fp.WriteLine(str);
                CreateAddTab(fp, 3); fp.WriteLine("}");//endfor
                                                       //GetReadDataByType
            }
            else if (datatype.m_Name.ToLower() == "enum")
            {
                var value = GetReadDataByType(datatype.m_Name, datatype.m_Data);//value
                var str = $"{TitleName} = {value}";
                CreateAddTab(fp, 3); fp.WriteLine(str);
            }
            else
            {
                var value = GetReadDataByType(datatype.m_Name, "");//value
                var str = $"{TitleName} = {value}";
                CreateAddTab(fp, 3); fp.WriteLine(str);
            }
        }
        CreateAddTab(fp, 2); fp.WriteLine("}");//end class
        CreateAddTab(fp, 1); fp.WriteLine("}");//start class
        fp.WriteLine("}");//start namespace
        fp.Close();
    }

    /// <summary>
    //REVIEW 엑셀 이름을 넣으면 데이터파일과 코드를 만들어 준다.
    /// </summary>
    /// <param name="xlsxName">엑셀 파일 풀네임</param>
    static void CreateExcelDataOne(string xlsxName)
    {
        var sheetNames = new List<string>();
        using (FileStream stream = File.Open(xlsxName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        {
            IWorkbook book = null;
            if (Path.GetExtension(xlsxName) == ".xls") book = new HSSFWorkbook(stream);
            else book = new XSSFWorkbook(stream);
            // 엑셀 파일 하단에 있는 sheet갯수만큼 데이터를 생성한다.
            for (int i = 0; i < book.NumberOfSheets; i++)
            {
                CExcelData _ExcelData = new CExcelData();

                int keyidx = -1;
                int typeidx = -1;
                List<int> rowList = _ExcelData.rowList;
                List<int> columnList = _ExcelData.columnList;
                List<string> typeList = new List<string>();
                List<string> KeyList = _ExcelData.m_ValueNames;
                var sheet = book.GetSheetAt(i);
                //시트 이름에 스킵 문자가 있으면 이건 스킵 해주자.
                if (sheet.SheetName.Contains(ExcelLoader.I.스킵표시문자)) continue;
                //엑셀의 A1~AN까지의 값을 읽어온다.
                for (int a = 0; a <= sheet.LastRowNum; a++)
                {
                    var row = sheet.GetRow(a);
                    if (row == null)
                    {
                        break;
                    }
                    //빈값이 있으면 0을 체워주자.
                    var firststr = row.GetCell(0)?.ToString() ?? "0";
                    //스킵 문자가 있으면 이 줄은 안읽는다.
                    if (firststr.Contains(ExcelLoader.I.스킵표시문자)) continue;
                    //데이터타입 관련 문자가 있으면 이건 데이터 타입이다.
                    if (firststr.Contains(ExcelLoader.I.데이터타입표시문자))
                    {
                        // 데이터 타입 문자가 있을때 첫번째이면 값을 저장한다.
                        if (typeidx < 0)
                        {
                            typeidx = a;
                        }
                        continue;
                    }
                    // 스킵도 아니고 데이터 타입도 아닐경우 처음 온 제대로된 데이터는 키값이 된다.
                    if (keyidx < 0)
                    {
                        //A{keyidx}~Z{keyidx}발향으로 읽어 가면서 데이터 값을 확인 해간다.
                        for (int b = 0; b < row.LastCellNum; b++)
                        {
                            //값을 문자열로 읽는다. 만약 셀값이 널이면 문자열을 빈값으로 채워준다.
                            var str = row.GetCell(b)?.ToString() ?? "";
                            //키값중에서도 스킵하고 싶은 데이터가 있을것이다. 그건 넘어가준다.
                            if (str.Contains(ExcelLoader.I.스킵표시문자)) continue;
                            //데이터들을 채워준다.
                            columnList.Add(b);
                            KeyList.Add(str);
                        }
                        keyidx = a;
                        continue;
                    }
                    rowList.Add(a);
                }
                // 타입을 따로 명시하지 않았을때는 전부다 문자열로 읽는다.
                if (typeidx < 0)
                {
                    foreach (var item in KeyList)
                    {
                        typeList.Add("string");
                    }
                }
                else
                {
                    foreach (var item in columnList)
                    {
                        //데이터 타입에 빈칸이 있으면 에러가 발생한다. 그럴경우 string로 값을 인식해준다.
                        try
                        {
                            var str = sheet.GetRow(typeidx).GetCell(item)?.ToString().Replace(" ", "").Replace(ExcelLoader.I.데이터타입표시문자, "") ?? "string";
                            typeList.Add(str);
                        }
                        catch
                        {
                            typeList.Add("string");
                        }
                    }
                }
                Dictionary<string, List<string>> enumDic = _ExcelData.m_EnumDic;
                //enum을 찾아서 데이터를 만들어 주자.
                for (int t = 0; t < typeList.Count; t++)
                {
                    var type = typeList[t];
                    //enum을 찾았다.
                    if (type.Contains("<") && type.Contains(">"))
                    {
                        //이넘형은 <>안에 데이터 형이 들어 간다. 그러기 위해서 "<",">"삭제 해주고 혹시 처음값이면 데이터형을 나타내는 문자도 삭제 해줘야 한다.
                        type = type.Replace(ExcelLoader.I.데이터타입표시문자, "").Replace("<", "").Replace(">", "");
                        //메모리 할당 안되어 있으면 할당해주고
                        if (!enumDic.ContainsKey(type))
                        {
                            enumDic[type] = new List<string>();
                        }
                        //이제 리스트에 넣어 주자.
                        var col = columnList[t];
                        for (int r = 0; r < rowList.Count; r++)
                        {
                            var str = sheet.GetRow(rowList[r]).GetCell(col).ToString();// saRet[rowList[r], col].ToString();
                            if (enumDic[type].Contains(str)) continue;
                            enumDic[type].Add(str);
                        }
                    }
                }
                List<CType> TypeList = _ExcelData.m_DataTypes;
                //데이터 형을 우선 만든다.
                foreach (var type in typeList)
                {
                    TypeList.Add(new CType(type, ExcelLoader.I.데이터타입표시문자));
                }
                // 파일 이름이 겹치는지를 확인하여 숫자를 붙인다음 진짜 사용할 파일 이름을 가져오자.
                var name = GetRealFileName(sheet.SheetName);
                var csfileName = name;
                if (ExcelLoader.I.cs파일이름소문자)
                {
                    csfileName = csfileName.ToLower();
                }
                if (ExcelLoader.I.cs첫글자대문자로)
                {
                    var first = csfileName[0];
                    csfileName = csfileName.Remove(0, 1);
                    csfileName = first.ToString().ToUpper() + csfileName;
                }
                csfileName = ExcelLoader.I.cs파일접두어 + csfileName + ExcelLoader.I.cs파일접미어;
                _ExcelData.m_Name = csfileName;
                _ExcelData.m_CSPath = Path.Combine(ExcelLoader.CODEPATH, csfileName) + ".cs";
                string path = Path.Combine(ExcelLoader.DOWNLOADPATH, name + ".bytes");
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
                #region bytes파일 저장
                System.IO.BinaryWriter fp = new BinaryWriter(System.IO.File.Open(path, FileMode.Create));
                // fp.Write(columnList.Count); //가로가 몇칸인가?
                fp.Write(rowList.Count);//세로가 몇줄인가?
                foreach (var row in rowList) // 한줄식 데이터를 저장하자.
                {
                    for (int c = 0; c < columnList.Count; c++)
                    {
                        var col = columnList[c];
                        var str = "0"; //sheet.GetRow(row).GetCell(col)?.ToString() ?? "0";
                        //빈칸인지 비교해서 빈칸이면 0을 넣어 주자.
                        if (sheet.GetRow(row) != null && sheet.GetRow(row).GetCell(col) != null && sheet.GetRow(row).GetCell(col).ToString().Replace(" ", "").Length > 0)
                        {
                            str = sheet.GetRow(row).GetCell(col).ToString();
                        }
                        else if (TypeList[c].m_Name.ToLower() == "string")//값이 없는데 스트링형이면 빈값을 넣어 주자.
                        {
                            str = "";
                        }
                        _Write(fp, TypeList[c], str, enumDic);
                    }
                }
                fp.Close();
                #endregion
                // System.IO.BinaryWriter fp = new System.IO.BinaryWriter(System.IO.File.Open(path, System.IO.FileMode.Create));
                // fp.Close() ;

                //데이터를 저장 해주자.
                sheetNames.Add(sheet.SheetName);
                m_CsDic[_ExcelData.m_Name] = name;
                CreateCode(_ExcelData);
                //  var name = GetRealFileName(sheet.SheetName);

            }
        }
    }
}
public class CExcelLoaderFactory
{
    //REVIEW - 단축키 % : CTRL  # : SHIFT & : ALT
    [MenuItem("Utility/Excel  Load %#E", false, 10)]
    static ExcelLoader CreateExcelLoader()
    {
        Selection.activeObject = ExcelLoader.I;
        return ExcelLoader.I;
    }
}