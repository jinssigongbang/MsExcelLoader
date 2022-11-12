using System;
using System.Collections.Generic;
using System.IO;
using UnityEditor;
using UnityEngine;

/// <summary>
//NOTE - 만든이 : 백수진모씨
//NOTE - 소속 : 15pluto
//NOTE - 만든날짜 : 2022-11.12
//NOTE - 문의 : yakmo94@naver.com
/// </summary>
#if UNITY_EDITOR
[InitializeOnLoad]
#endif
[Serializable]
public class ExcelLoader : ScriptableObject
{
    [HideInInspector]
    public List<string> m_Paths = new List<string>() ;
    [SerializeField]
    [Tooltip("로딩해서 데이터로 만들 엑셀 파일들")]
    private UnityEngine.Object [] 엑셀파일들;
    [SerializeField]
    [Tooltip("데이터를 저장할 위치 /Resources나 하위 폴더만 가능하다")]
    private UnityEngine.Object bytes저장위치;
    [SerializeField]
    [Tooltip("코드가 생성될 위치")]
    private UnityEngine.Object 코드생성위치;
    [Header("cs파일 관련 옵션")]
    [Tooltip(".cs파일을 생성할때 맨앞에 붙일 단어")]
    public string cs파일접두어 = "";
    [Tooltip(".cs파일을 생성할때 맨뒤에 붙일 단어")]
    public string cs파일접미어 = "CExcel";
   
    [Tooltip(".cs파일을 이름의 첫글자를 대문자로 할것인가?")]
    public bool cs첫글자대문자로 = false; 
    [Tooltip(".cs파일을 이름을 소문자로 할것인가?")]
    public bool cs파일이름소문자 = false;
    [Tooltip("이 문자가 포함되면 sheet나 데이터를 읽지 않는다.")]
    public string 스킵표시문자 = "@";
    [Tooltip("row의 시작에 이 문자가 포함 되면 그줄은 데이터 타입으로 인식한다")]
    public string 데이터타입표시문자 = "#";
     [Tooltip("이 이름으로 클레스를 만든다.")]
    public string Manager클레스이름 = "CMsExcelManager";
    public UnityEngine.Object[] ExcelDatas
    {
        get => 엑셀파일들;
    }
    #if UNITY_EDITOR
    public static string DOWNLOADPATH
    {
        get => I.bytes저장위치 ? AssetDatabase.GetAssetPath(I.bytes저장위치) : "";
    }
    
    public static string RESOURCEPATH
    {
        get
        {
            var idx = AssetDatabase.GetAssetPath(I.bytes저장위치).IndexOf("Resources");
            return I.bytes저장위치 ? AssetDatabase.GetAssetPath(I.bytes저장위치).Remove(0, idx) : "";
        }
    }
    public static string CODEPATH
    {
        get=>I.코드생성위치 ? AssetDatabase.GetAssetPath(I.코드생성위치) : "";
    }
    #endif
    public static int ExcelCount
    {
        get=>I.엑셀파일들.Length ;
    }

    const string configAssetName = "_ExcelLoader";
    private static ExcelLoader m_Instance;
    public static ExcelLoader I
    {
        get
        {
            if (m_Instance == null)
            {
                m_Instance = Resources.Load(configAssetName) as ExcelLoader;
                if (m_Instance == null)
                {
                    m_Instance = CreateInstance<ExcelLoader>();
#if UNITY_EDITOR
                    if (!System.IO.Directory.Exists(Path.Combine(Application.dataPath, "Resources")))
                    {
                        var dir = new System.IO.DirectoryInfo(Application.dataPath);
                        dir.CreateSubdirectory("Resources");
                    }
                    AssetDatabase.CreateAsset(m_Instance, $"Assets/Resources/{configAssetName}.asset");
#endif
                }
            }
            return m_Instance;
        }
    }
}
