using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
/// <summary>
//NOTE - 만든이 : 백수진모씨
//NOTE - 소속 : 15pluto
//NOTE - 만든날짜 : 2022-11.12
//NOTE - 문의 : yakmo94@naver.com
/// </summary>
namespace ExcelConverter
{
    public class CType
    {
        public string m_Name = "";
        public string m_Data = "";
        public CType(string type,string dataTypeSeparator)
        {
            type = type.Replace(dataTypeSeparator, "");
            if (type.Contains("<"))
            {
                m_Name = "enum";
                type = type.Replace(dataTypeSeparator, "").Replace("<", "").Replace(">", "");
                //이넘형은 이름만 넣는다.
                m_Data = type.Split(',')[0];
            }
            //List
            else if (type.Contains("["))
            {
                m_Name = "List";
                type = type.Replace(dataTypeSeparator, "").Replace("[", "").Replace("]", "");
                var items = type.Split(',');
                m_Data = GetType(items[0]);
            }
            else
            {
                m_Name = GetType(type);
            }
        }
        public string GetType(string str)
        {
            str = str.Replace(" ","").ToLower();
            if (str.Equals("boolean") || str.Equals("bool")) return "bool";
            if (str.Equals("byte") || str.Equals("uint8")) return "byte";
            if (str.Equals("char") || str.Equals("int8")) return "short";
            if (str.Equals("short") || str.Equals("int16")) return "short";
            if (str.Equals("ushort") || str.Equals("uint16")) return "UInt16";
            if (str.Equals("int") || str.Equals("int32")) return "int";
            if (str.Equals("uint") || str.Equals("uint32")) return "UInt32";
            if (str.Equals("int64") || str.Equals("long")) return "Int64";
            if (str.Equals("uint64") || str.Equals("ulong")) return "UInt64";
            if (str.Equals("float")) return "float";
            if (str.Equals("double")) return "double";
            return str;
        }
    }
    public class CExcelData
    {
        public string m_Name;
        public string m_CSPath;
        public List<string> m_ValueNames = new List<string>();
        public List<string> m_Items = new List<string>();
        public List<CType> m_DataTypes = new List<CType>();
        public List<int> rowList = new List<int>();
        public List<int> columnList = new List<int>();
        public Dictionary<string, List<string>> m_EnumDic = new Dictionary<string, List<string>>();
    }
}
