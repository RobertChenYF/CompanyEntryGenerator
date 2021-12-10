using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using OfficeOpenXml;
using System.IO;

public class GenerateEntryBook : MonoBehaviour
{
    private string EntryBookPath;
    private string AllEntryContenPath;
    public Excel entryBook;
    public Excel AllEntryContentBook;
    public List<EntryContent> AllEntryContent;
    public List<EntryContent> AbnormalEntryContent;
    public int TotalEntryNeeded = 20;
    public int TotalAbnormalEntry = 5;
    public List<string> NameList;
    public List<string> BadNameList;
    public List<string> UnusedNameList;
    public string BadNameBuffer = null;

    private List<string> SuspiciousEntry;


    // Start is called before the first frame update
    void Start()
    {
        SuspiciousEntry = new List<string> { "Suspicious Amount", "Suspicious Price","Suspicious Name", "Invalid Employee" };
        Debug.Log(SuspiciousEntry.Count);
        AllEntryContent = new List<EntryContent>();
        AbnormalEntryContent = new List<EntryContent>();
        NameList = new List<string>();
        UnusedNameList = new List<string>();
        BadNameList = new List<string>();

        EntryBookPath = Application.dataPath + "/CompanyEntryBook.xlsx";
        AllEntryContenPath = Application.dataPath + "/EntryBookContent.xlsx";
        
        entryBook = ExcelHelper.LoadExcel(EntryBookPath);
        AllEntryContentBook = ExcelHelper.LoadExcel(AllEntryContenPath);
        ReadNameList();
        ReadEntryContentList();
        ReadAbnormalEntryContentList();
        //Debug.Log(AllEntryContent.Count);

    }

    // Update is called once per frame
    void Update()
    {
        if (Input.GetKeyDown(KeyCode.Space))
        {
            GenerateEntry(TotalEntryNeeded, TotalAbnormalEntry);
        }
    }

    public void ReadNameList()
    {
        int a = 2;
        string b;
        while (a<30)
        {
            b = AllEntryContentBook.Tables[1].GetValue(a, 1).ToString();
            //Debug.Log(b);
            NameList.Add(b);
            UnusedNameList.Add(b);
            a++;
        }
        while (a<40)
        {
            b = AllEntryContentBook.Tables[1].GetValue(a, 1).ToString();
            
            UnusedNameList.Add(b);
            a++;
        }

    }

    public void ReadEntryContentList()
    {
        int a = 2;
        while ((AllEntryContentBook.Tables[0].GetValue(a,1).ToString().Length > 0))
        {
            EntryContent entry = new EntryContent(AllEntryContentBook.Tables[0].GetValue(a, 1).ToString(), int.Parse(AllEntryContentBook.Tables[0].GetValue(a, 4).ToString()),
               int.Parse(AllEntryContentBook.Tables[0].GetValue(a, 5).ToString()), float.Parse(AllEntryContentBook.Tables[0].GetValue(a, 2).ToString()),
               float.Parse(AllEntryContentBook.Tables[0].GetValue(a, 3).ToString()), int.Parse(AllEntryContentBook.Tables[0].GetValue(a, 6).ToString()));

            AllEntryContent.Add(entry);
            a++;
        }
    }
    public void ReadAbnormalEntryContentList()
    {
        int a = 2;
        while ((AllEntryContentBook.Tables[2].GetValue(a, 1).ToString().Length > 0))
        {
            EntryContent entry = new EntryContent(AllEntryContentBook.Tables[2].GetValue(a, 1).ToString(), int.Parse(AllEntryContentBook.Tables[2].GetValue(a, 4).ToString()),
               int.Parse(AllEntryContentBook.Tables[2].GetValue(a, 5).ToString()), float.Parse(AllEntryContentBook.Tables[2].GetValue(a, 2).ToString()),
               float.Parse(AllEntryContentBook.Tables[2].GetValue(a, 3).ToString()), int.Parse(AllEntryContentBook.Tables[2].GetValue(a, 6).ToString()));

            AbnormalEntryContent.Add(entry);
            a++;
        }
    }

    public void WriteToEntryBook(int sheetNumber,int column, int row, string content)
    {
        entryBook.Tables[sheetNumber].SetValue(row,column,content);
    }

    public string GenerateAGoodName()
    {
        int a = Random.Range(0, NameList.Count);
       // Debug.Log(a);
        string name = NameList[a];
        if (UnusedNameList.Contains(name))
        {
            UnusedNameList.Remove(name);
        }

        return name;
    }

    public string GenerateABadName()
    {

        string name;

            int a = Random.Range(0, UnusedNameList.Count);
            // Debug.Log(a);
            name = UnusedNameList[a];
            if (NameList.Contains(name))
            {
                NameList.Remove(name);
            }

            UnusedNameList.Remove(name);
            BadNameList.Add(name);
        

        return name;
    }

    
    public void GenerateEntry(int totalAmount,int totalAbnormalAmount)
    {
        List<int> AbnormalEntryRow = new List<int>();
        for (int i = totalAbnormalAmount; i > 0; i --)
        {
            int RandomRow = Random.Range(2, 2 + totalAmount);
            while (AbnormalEntryRow.Contains(RandomRow))
            {
                RandomRow = Random.Range(2, 2 + totalAmount);
            }
            AbnormalEntryRow.Add(RandomRow);
        }
        int abnormalCount = 2;
        for (int i = 2; i <= 2 + totalAmount; i++ )
        {
            if (!AbnormalEntryRow.Contains(i))
            {
                //Debug.Log(i);
                GenerateOneNormalEntry(i);
            }
            else
            {
                GenrateOneAbnormalEntry(i,abnormalCount);
                abnormalCount++;
            }
        }

        ExcelHelper.SaveExcel(entryBook, EntryBookPath);

    }

    public void GenerateOneNormalEntry(int row)
    {
        int a = Random.Range(0, AllEntryContent.Count);
        //Debug.Log(a);
        //write to entry content
        //Debug.Log(AllEntryContent[a].EntryName);
        WriteToEntryBook(0, 1, row, AllEntryContent[a].EntryName);
        int Amount =Mathf.FloorToInt(Random.Range(0.0f,1.0f)*Random.Range(AllEntryContent[a].highestAmount-AllEntryContent[a].normalAmount,0)) + AllEntryContent[a].normalAmount ;
        WriteToEntryBook(0, 2, row,Amount.ToString());
        float averagePrice = Mathf.Floor(Random.Range(0.0f, 1.0f) * Random.Range(AllEntryContent[a].highestPrice - AllEntryContent[a].normalPrice, 0)*100)/100.0f + AllEntryContent[a].normalPrice;
        WriteToEntryBook(0, 3, row, (averagePrice * Amount).ToString());
        string c = GenerateAGoodName();
//Debug.Log(c);
        WriteToEntryBook(0, 4, row, c);


        AllEntryContent[a].AmountAppear--;
        if(AllEntryContent[a].AmountAppear == 0)
        {
            AllEntryContent.RemoveAt(a);
        }
    }

    public void GenrateOneAbnormalEntry(int row, int count)
    {
        int a;
        //check which abnormal case
        if (count == 2)
        {
            a = Random.Range(0, SuspiciousEntry.Count - 1);
        }
        else
        {
            a = Random.Range(0, SuspiciousEntry.Count);
            //Debug.Log(SuspiciousEntry.Count);
            
        }

        

        //write to the entrybook
        if(a == 0)
        {
            GenerateAbnormalAmountEntry(row);

        }
        else if (a == 1)
        {
            GenerateAbnormalSpendingEntry(row);
        }
        else if (a == 2)
        {
            GenerateSuspiciousEntryName(row);
        }
        else if (a == 3)
        {
            GenerateBadNameEntry(row);
        }
        //write to the answersheet


        WriteToEntryBook(2,1,count,row.ToString());
        WriteToEntryBook(2, 2, count, SuspiciousEntry[a]);
        

    }

    public void GenerateAbnormalSpendingEntry(int row)
    {
        int a = Random.Range(0, AllEntryContent.Count);
       
        WriteToEntryBook(0, 1, row, AllEntryContent[a].EntryName);
        int Amount = Mathf.FloorToInt(Random.Range(0.0f, 1.0f) * Random.Range(AllEntryContent[a].highestAmount - AllEntryContent[a].normalAmount, 0)) + AllEntryContent[a].normalAmount;
        WriteToEntryBook(0, 2, row, Amount.ToString());
        float Price = Mathf.Floor(Random.Range(1.0f, 3.0f) *(AllEntryContent[a].highestPrice - AllEntryContent[a].normalPrice) * 100) / 100.0f + AllEntryContent[a].highestPrice;
        WriteToEntryBook(0, 3, row, (Price * Amount).ToString());
        WriteToEntryBook(0, 4, row, GenerateABadName());
        AllEntryContent[a].AmountAppear--;
        if (AllEntryContent[a].AmountAppear == 0)
        {
            AllEntryContent.RemoveAt(a);
        }
    }

    public void GenerateSuspiciousEntryName(int row)
    {
        int a = Random.Range(0, AbnormalEntryContent.Count);
        //Debug.Log(a);
       
        WriteToEntryBook(0, 1, row, AbnormalEntryContent[a].EntryName);
        int Amount = Mathf.FloorToInt(Random.Range(0.0f, 1.0f) * Random.Range(AbnormalEntryContent[a].highestAmount - AbnormalEntryContent[a].normalAmount, 0)) + AbnormalEntryContent[a].normalAmount;
        WriteToEntryBook(0, 2, row, Amount.ToString());
        float averagePrice = Mathf.Floor(Random.Range(0.0f, 1.0f) * Random.Range(AbnormalEntryContent[a].highestPrice - AbnormalEntryContent[a].normalPrice, 0) * 100) / 100.0f + AbnormalEntryContent[a].normalPrice;
        WriteToEntryBook(0, 3, row, (averagePrice * Amount).ToString());
        string c = GenerateABadName();
        //Debug.Log(c);
        WriteToEntryBook(0, 4, row, c);


        AbnormalEntryContent[a].AmountAppear--;
        if (AbnormalEntryContent[a].AmountAppear == 0)
        {
            AbnormalEntryContent.RemoveAt(a);
        }
    }

    public void GenerateAbnormalAmountEntry(int row)
    {
        int a = Random.Range(0, AllEntryContent.Count);
        //write to entry content
        
        WriteToEntryBook(0, 1, row, AllEntryContent[a].EntryName);
        int Amount = Mathf.RoundToInt(Random.Range(2.0f, 4.0f) *(AllEntryContent[a].highestAmount)) + AllEntryContent[a].highestAmount;
        WriteToEntryBook(0, 2, row, Amount.ToString());
        float averagePrice = Mathf.Floor(Random.Range(0.0f, 1.0f) * Random.Range(AllEntryContent[a].highestPrice - AllEntryContent[a].normalPrice, 0) * 100) / 100.0f + AllEntryContent[a].normalPrice;
        WriteToEntryBook(0, 3, row, (averagePrice * Amount).ToString());
        WriteToEntryBook(0, 4, row, GenerateABadName());


        AllEntryContent[a].AmountAppear--;
        if (AllEntryContent[a].AmountAppear == 0)
        {
            AllEntryContent.RemoveAt(a);
        }
    }

    public void GenerateBadNameEntry(int row)
    {
        int a = Random.Range(0, AllEntryContent.Count);
        //Debug.Log(a);
        //write to entry content
        //Debug.Log(AllEntryContent[a].EntryName);
        WriteToEntryBook(0, 1, row, AllEntryContent[a].EntryName);
        int Amount = Mathf.FloorToInt(Random.Range(0.0f, 1.0f) * Random.Range(AllEntryContent[a].highestAmount - AllEntryContent[a].normalAmount, 0)) + AllEntryContent[a].normalAmount;
        WriteToEntryBook(0, 2, row, Amount.ToString());
        float averagePrice = Mathf.Floor(Random.Range(0.0f, 1.0f) * Random.Range(AllEntryContent[a].highestPrice - AllEntryContent[a].normalPrice, 0) * 100) / 100.0f + AllEntryContent[a].normalPrice;
        WriteToEntryBook(0, 3, row, (averagePrice * Amount).ToString());


            
            string c = BadNameList[Random.Range(0,BadNameList.Count)];
            
            WriteToEntryBook(0, 4, row, c);


            AllEntryContent[a].AmountAppear--;
            if (AllEntryContent[a].AmountAppear == 0)
            {
                AllEntryContent.RemoveAt(a);
            }
        
    }
}

 public class EntryContent
{
    public string EntryName;
    public int normalAmount;
    public int highestAmount;
    public float normalPrice;
    public float highestPrice;
    public int AmountAppear;

    public EntryContent(string Name, int nAmount, int hAmount, float nPrice, float hPrice, int Appear)
    {
        EntryName = Name;
        normalAmount = nAmount;
        highestAmount = hAmount;
        normalPrice = nPrice;
        highestPrice = hPrice;
        AmountAppear = Appear;
     
     }


}
