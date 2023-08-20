using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace SapShellDesign
{
    [Serializable]
    public struct DesignDefinition
    {
        public int DesignMethod;
        public double Fy;
        public double Fs;
        public double Fc;
        public int TempRein;
        public double PosCover;
        public double NegCover;
        public int PosRowNum;
        public int NegRowNum;
        public double MaxReinSpace;
    }
    [Serializable]
    public struct OutputDefinition
    {
        public int Combo;
        public int StripMomentType;
        public int StripOutputType;
        public int[] CrackCombos;
        public int[] StrengthCombos;

    }
    [Serializable]
    public struct SelectionData
    {
        public string StuctureName;
        public string ShellName;
        public string StripName;
        public DesignDefinition DesignDefinition;
        public OutputDefinition OutputDefinition;
        public string SapFilePath;
        public List<string> SelElementNames;
    }
    class HistoryMaker
    {
        public static void SaveSelectionToFile(string fPath, List<SelectionData> selData)
        {
            string strData = GenericSerializer.Serialize(selData);
            var streamWriter = new StreamWriter(fPath, false);
            streamWriter.Write(strData);
            streamWriter.Close();
        }
        public static List<SelectionData> LoadData(string fPath)
        {
            var streamReader = new StreamReader(fPath);
            string strData = streamReader.ReadToEnd();
            streamReader.Close();
            var selData = new List<SelectionData>();
            if (GenericSerializer.DeSerialize(strData, ref selData))
            {
                return selData;
            }
            return new List<SelectionData>();
        }
        //private static void OpenFile(StructFile sFile)
        //{
        //    if (sFile != null)
        //    {
        //        sFile.Close();
        //        sFile = null;
        //    }
        //    sFile.Open(FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        //}

        //public static void AddSelectionToFile(string fPath, SelectionData selData)
        //{

            //var sf = new StructFile(fPath, typeof(SelectionData));
            //sf.Open(FileMode.Append,FileAccess.Write,FileShare.None);
            //sf.WriteStructure(selData);
            //sf.Close();
        //}
        //public static SelectionData LoadData(StructFile sFile)
        //{
        //    if (sFile == null) OpenFile(sFile);
        //    return (SelectionData)sFile.GetNextStructureValue();
        //}
        //private static void OpenFile(StructFile sFile)
        //{
        //    if (sFile != null)
        //    {
        //        sFile.Close();
        //        sFile = null;
        //    }
        //    sFile.Open(FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        //}
    }
}
