using System;
[Serializable]
public class RechargeData
{
    public int Coding;
    public string Icon;
    public string Lable;
    public int Before;
    public int Now;
    public string Rate;
    public float Price;
}
[Serializable]
public class PropData
{
    public int Coding;
    public string Icon;
    public string Lable;
    public string Name;
    public string Introduce;
    public float Price;
}
[Serializable]
public class GiftPackageData
{
    public int Coding;
    public string Name;
    public int Kind;
    public string OneIcon;
    public int OneCount;
    public string TwoIcon;
    public int TwoCount;
    public string ThreeIcon;
    public int ThreeCount;
    public float Price;
}
