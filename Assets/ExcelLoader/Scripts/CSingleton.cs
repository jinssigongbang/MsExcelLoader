using System.Collections;
using System.Collections.Generic;
using UnityEngine;
/// <summary>
//NOTE - 만든이 : 백수진모씨
//NOTE - 소속 : 15pluto
//NOTE - 만든날짜 : 2022-11.12
//NOTE - 문의 : yakmo94@naver.com
/// </summary>
public class CSingletonMono<T> : MonoBehaviour where T : MonoBehaviour
{
    static T m_Instance ;
    public static T I
    {
        get
        {
            if(m_Instance == null)
            {
                m_Instance = GameObject.FindObjectOfType<T>() ;
                if(m_Instance == null)
                {
                    m_Instance = new GameObject(typeof(T).ToString() + "_(singleton)").AddComponent<T>() ;
                    DontDestroyOnLoad(m_Instance.gameObject) ;
                }
            }
            return m_Instance ;
        }
    }
}
public class CSingleton<T> where T : class , new()
{
    static T m_Instance = new T() ;
    public static T I => m_Instance ;
    public virtual void Init(){}
    public string Name
    {
        get => typeof(T).ToString();
    }
}
