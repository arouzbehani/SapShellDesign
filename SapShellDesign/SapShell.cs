using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using SAP2000v1;

namespace SapShellDesign
{
    [Serializable]
    public struct ShellElement
    {
        public string Name;
        internal Dictionary<string, Point> Points;
        internal ShellDirection PlaneDirection;
        private Dictionary<string, Point> xSortedPoints;
        private Dictionary<string, Point> ySortedPoints;
        private Dictionary<string, Point> zSortedPoints;
        public double L11;
        public double L22;
        public double LengthX;
        public double LengthY;
        public double LengthZ;
        public void SetPoints(Dictionary<string, Point> points)
        {
            Points = points;
            xSortedPoints = SortedByX();
            ySortedPoints = SortedByY();
            zSortedPoints = SortedByZ();
            CalculateLengths();
        }
        public void SetDirection(ShellDirection planeDirection)
        {
            PlaneDirection = planeDirection;
        }

        private Dictionary<string, Point> SortedByX()
        {
            return SapShell.GetMyDic(Points.OrderBy(point => point.Value.X));
        }
        private Dictionary<string, Point> SortedByY()
        {
            return SapShell.GetMyDic(Points.OrderBy(point => point.Value.Y));
        }
        private Dictionary<string, Point> SortedByZ()
        {
            return SapShell.GetMyDic(Points.OrderBy(point => point.Value.Z));
        }
        private void CalculateLengths()
        {
            if (PlaneDirection == ShellDirection.XY || PlaneDirection == ShellDirection.XZ)
            {
                var myList = xSortedPoints.Values.ToList();
                LengthX = myList[myList.Count - 1].X - myList[0].X;
            }
            else
            {
                LengthX = 0;
            }

            if (PlaneDirection == ShellDirection.XY || PlaneDirection == ShellDirection.YZ)
            {
                var myList = ySortedPoints.Values.ToList();
                LengthY = myList[myList.Count - 1].Y - myList[0].Y;
            }
            else
            {
                LengthY = 0;
            }

            if (PlaneDirection == ShellDirection.XZ || PlaneDirection == ShellDirection.YZ)
            {
                var myList = zSortedPoints.Values.ToList();
                LengthZ = myList[myList.Count - 1].Z - myList[0].Z;
            }
            else
            {
                LengthZ = 0;
            }
        }
        public bool ContainsPoint(Point point)
        {
            if (Points.ContainsValue(point))
            {
                return true;
            }
            return false;
        }
        public Vector GetNormalVector()
        {
            Point p0 = Points.Values.ToList()[0];
            Point p1 = Points.Values.ToList()[1];
            Point p2 = Points.Values.ToList()[2];

            var v01 = new Vector() { Pi = p0.XYZ(), Pj = p1.XYZ() };
            var v02 = new Vector() { Pi = p0.XYZ(), Pj = p2.XYZ() };

            return VectorOperations.GetNormalVector(v01, v02);

        }
        public Plane GetPlane()
        {
            Point p0 = Points.Values.ToList()[0];
            Point p1 = Points.Values.ToList()[1];
            Point p2 = Points.Values.ToList()[2];
            return new Plane(p0.XYZ(), p1.XYZ(), p2.XYZ());
        }
    }

    [Serializable]
    public struct Point
    {
        public string Name;
        public double X;
        public double Y;
        public double Z;
        public Dictionary<string, double[]> Loads;
        public XYZPoint XYZ()
        {
            return new XYZPoint() { X = X, Y = Y, Z = Z };
        }
    }
    [Serializable]
    public enum ShellDirection
    {
        XY,
        XZ,
        YZ,
        Other
    }
    [Serializable]
    public struct ContourPos
    {
        public double X;
        public double Y;
        public double Z;
        public double Value;
    }

    public class SapShell
    {
        readonly string ProgramPath = "C:\\Program Files\\Computers and Structures\\SAP2000 22\\SAP2000.exe";
        private int ret;
        private List<ShellElement> ElementCol;
        private Dictionary<string, List<ShellElement>> PointCol;
        private Dictionary<string, Point> AllShellPoints;
        private cOAPI SapObj;
        private cSapModel SapModel;
        public string SapFilePath;
        private bool SapStarted;
        cHelper helper;
        public SapShell(string filePath)
        {
            SapFilePath = filePath;
            StartSap();
        }
        private void StartSap()
        {
            helper = new Helper();
            SapObj = helper.CreateObject(ProgramPath);
            //SapObj = (cOAPI)System.Runtime.InteropServices.Marshal.GetActiveObject("CSI.SAP2000.API.SapObject");
            if (string.IsNullOrEmpty(SapFilePath))
            {
                MessageBox.Show("File Path Missed!", "SAP2000 Shell Designer");
                SapStarted = false;
            }
            SapObj.ApplicationStart(eUnits.Ton_m_C, true, SapFilePath);
            SapModel = SapObj.SapModel;
            ret = SapModel.SetModelIsLocked(false);
            SapStarted = true;
        }
        private void EndSap()
        {
            SapObj.ApplicationExit(true);

            //catch (Exception)
            //{
            KillProcesses("Sap2000");
            //}

        }
        private void ReOpenSap(string fPath)
        {
            Thread.Sleep(3333);
            SapObj = helper.CreateObject(ProgramPath);
            SapObj.ApplicationStart(eUnits.Ton_m_C, true, fPath);
            SapModel = SapObj.SapModel;
            ret = SapModel.SetModelIsLocked(false);
        }

        private void CalculateShells()
        {
            var myList = new List<ShellElement>();
            if (SapStarted)
            {
                int objNum = 0;
                string[] objNames = null;
                ret = SapModel.AreaObj.GetNameList(ref objNum, ref objNames);
                for (int i = 0; i < objNum; i++)
                {
                    myList.Add(GetMyElement(objNames.GetValue(i).ToString()));
                }
                ElementCol = myList;
                Run();
                EndSap();
            }
        }

        private void Run()
        {
            int caseNum = 0;
            string[] caseNames = null;
            ret = SapModel.LoadCases.GetNameList(ref caseNum, ref caseNames);

            int combNum = 0;
            string[] combNames = null;
            ret = SapModel.RespCombo.GetNameList(ref combNum, ref combNames);


            ret = SapModel.Analyze.SetRunCaseFlag("Dead", false, true);
            ret = SapModel.Analyze.SetRunCaseFlag("MODAL", false);
            for (int i = 0; i < caseNum; i++)
            {
                if (caseNames.GetValue(i).ToString().ToLower() != "modal")
                    ret = SapModel.Analyze.SetRunCaseFlag(caseNames.GetValue(i).ToString(), true);
            }
            ret = SapModel.Analyze.CreateAnalysisModel();
            ret = SapModel.Analyze.RunAnalysis();

            ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput();
            for (int i = 0; i < combNum; i++)
            {
                ret = SapModel.Results.Setup.SetComboSelectedForOutput(combNames.GetValue(i).ToString(), true);
            }
            int numberResult = 0;
            string[] obj = null;
            string[] elem = null;
            string[] pointElem = null;
            string[] loadCases = null;
            string[] stepType = null;
            double[] stepNum = null;
            double[] f11 = null;
            double[] f22 = null;
            double[] f12 = null;
            double[] fMax = null;
            double[] fMin = null;
            double[] fAngle = null;
            double[] fVM = null;
            double[] m11 = null;
            double[] m22 = null;
            double[] m12 = null;
            double[] mMax = null;
            double[] mMin = null;
            double[] mAngle = null;
            double[] v13 = null;
            double[] v23 = null;
            double[] vMax = null;
            double[] vAngle = null;

            PointCol = new Dictionary<string, List<ShellElement>>();
            AllShellPoints = new Dictionary<string, Point>();
            foreach (ShellElement element in ElementCol)
            {
                ret = SapModel.Results.AreaForceShell(element.Name, eItemTypeElm.ObjectElm, ref numberResult, ref obj, ref elem,
                                                      ref pointElem, ref loadCases, ref stepType, ref stepNum, ref f11, ref f22, ref f12, ref fMax, ref fMin,
                                                      ref fAngle, ref fVM, ref m11, ref m22, ref m12, ref mMax, ref mMin, ref mAngle, ref v13, ref v23, ref vMax,
                                                      ref vAngle);
                foreach (KeyValuePair<string, Point> point in element.Points)
                {
                    string pName = point.Key;

                    for (int i = 0; i < numberResult; i++)
                    {
                        if (pointElem.GetValue(i).ToString() == pName)
                        {
                            if (!AllShellPoints.ContainsKey(pName))
                            {
                                AllShellPoints.Add(pName, point.Value);
                            }
                            if (!PointCol.ContainsKey(pName))
                            {
                                PointCol.Add(pName, new List<ShellElement>() { element });
                            }
                            else
                            {
                                if (!PointCol[pName].Contains(element))
                                    PointCol[pName].Add(element);
                            }
                            if (!point.Value.Loads.ContainsKey(loadCases.GetValue(i).ToString()))
                            //if (!dic.ContainsKey(loadCases.GetValue(i).ToString()))
                            {
                                point.Value.Loads.Add(loadCases.GetValue(i).ToString(),
                                                  new[] { Convert.ToDouble(m11.GetValue(i)), Convert.ToDouble(m22.GetValue(i)) });
                                //dic.Add(loadCases.GetValue(i).ToString(),
                                //        new[] {Convert.ToDouble(m11.GetValue(i)), Convert.ToDouble(m22.GetValue(i))});
                            }
                        }
                    }
                    //element.SetPointLoads(dic,pName);
                }
            }
        }

        public void SelectElements(List<string> elementNames)
        {
            ret = SapModel.AreaObj.SetSelected("All", false, eItemType.Group);
            for (int i = 0; i < elementNames.Count; i++)
            {
                ret = SapModel.AreaObj.SetSelected(elementNames[i], true, eItemType.Objects);
            }
            ret = SapModel.View.RefreshWindow(0);
        }
        private ShellElement GetMyElement(string elementName)
        {
            if (elementName == "3246")
            {
                string s = "a";
            }
            var myELement = new ShellElement() { Name = elementName };
            int pn = 0;
            string[] points = null;
            double x = 0;
            double y = 0;
            double z = 0;
            ret = SapModel.AreaObj.GetPoints(elementName, ref pn, ref points);
            var xCol = new List<double>();
            var yCol = new List<double>();
            var zCol = new List<double>();
            var myPoints = new Dictionary<string, Point>();
            for (int i = 0; i < pn; i++)
            {
                string pName = points.GetValue(i).ToString();
                ret = SapModel.PointObj.GetCoordCartesian(pName, ref x, ref y, ref z);
                myPoints.Add(pName, new Point { Name = pName, X = Math.Round(x, 3), Y = Math.Round(y, 3), Z = Math.Round(z, 3), Loads = new Dictionary<string, double[]>() });
                if (!xCol.Contains(Math.Round(x, 3)))
                    xCol.Add(Math.Round(x, 3));
                if (!yCol.Contains(Math.Round(y, 3)))
                    yCol.Add(Math.Round(y, 3));
                if (!zCol.Contains(Math.Round(z, 3)))
                    zCol.Add(Math.Round(z, 3));
            }

            if (xCol.Count == 1)
            {
                myELement.SetDirection(ShellDirection.YZ);
            }
            if (yCol.Count == 1)
            {
                myELement.SetDirection(ShellDirection.XZ);
            }
            if (zCol.Count == 1)
            {
                myELement.SetDirection(ShellDirection.XY);
            }
            myELement.SetPoints(myPoints);

            return myELement;
        }
        private void KillProcesses(string processName)
        {
            Process[] Processes = Process.GetProcessesByName(processName);
            foreach (Process process in Processes)
            {
                process.Kill();
            }
        }
        public void GetShells(out Dictionary<string, ShellElement> allElements,
            out Dictionary<string, List<ShellElement>> xyPlanes,
            out Dictionary<string, List<ShellElement>> xzPlanes,
            out Dictionary<string, List<ShellElement>> yzPlanes,
            out string[] comboNames)
        {
            var xyElementsDic = new Dictionary<double, List<ShellElement>>();
            var xzElementsDic = new Dictionary<double, List<ShellElement>>();
            var yzElementsDic = new Dictionary<double, List<ShellElement>>();
            int combNum = 0;
            string[] combNames = null;
            ret = SapModel.RespCombo.GetNameList(ref combNum, ref combNames);
            //if (combNames.Any())
            {
                comboNames = new string[combNames.Length];
                for (int i = 0; i < combNames.Length; i++)
                {
                    comboNames[i] = (string)combNames.GetValue(i);
                }

            }
            xyPlanes = new Dictionary<string, List<ShellElement>>();
            xzPlanes = new Dictionary<string, List<ShellElement>>();
            yzPlanes = new Dictionary<string, List<ShellElement>>();
            allElements = new Dictionary<string, ShellElement>();
            CalculateShells();
            foreach (ShellElement element in ElementCol)
            {
                allElements.Add(element.Name, element);
            }
            try
            {
                var xyshells = from shellElement in ElementCol
                               where (shellElement.PlaneDirection == ShellDirection.XY)
                               select shellElement;
                var elements = new List<ShellElement>();
                foreach (ShellElement shellElement in xyshells)
                {
                    double z = shellElement.Points.Values.ToList()[0].Z;
                    if (!xyElementsDic.ContainsKey(z))
                    {
                        xyElementsDic.Add(z, new List<ShellElement>() { shellElement });
                    }
                    else
                    {
                        xyElementsDic[z].Add(shellElement);
                    }
                    Plane xyP = shellElement.GetPlane();
                    foreach (KeyValuePair<string, List<ShellElement>> xyPlane in xyPlanes)
                    {
                        if (VectorOperations.IsInPlane(xyP, xyPlane.Value[0].Points.Values.ToList()[0].XYZ()))
                        {

                        }
                    }

                    if (!xyPlanes.ContainsKey(xyP.GetPlaneScript()))
                    {
                        elements.Add(shellElement);
                        xyPlanes.Add(xyP.GetPlaneScript(), elements);
                    }
                    else
                    {
                        xyPlanes[xyP.GetPlaneScript()].Add(shellElement);
                    }
                }
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.ToString());
            }
            try
            {
                var xzshells = from shellElement in ElementCol
                               where shellElement.PlaneDirection == ShellDirection.XZ
                               select shellElement;
                var elements = new List<ShellElement>();
                foreach (ShellElement shellElement in xzshells)
                {
                    double y = shellElement.Points.Values.ToList()[0].Y;
                    if (!xzElementsDic.ContainsKey(y))
                    {
                        xzElementsDic.Add(y, new List<ShellElement>() { shellElement });
                    }
                    else
                    {
                        xzElementsDic[y].Add(shellElement);
                    }
                    Plane xzP = shellElement.GetPlane();
                    if (!xzPlanes.ContainsKey(xzP.GetPlaneScript()))
                    {
                        elements.Add(shellElement);
                        xzPlanes.Add(xzP.GetPlaneScript(), elements);
                    }
                    else
                    {
                        xzPlanes[xzP.GetPlaneScript()].Add(shellElement);
                    }
                }
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.ToString());
            }

            try
            {
                var yzshells = from shellElement in ElementCol
                               where shellElement.PlaneDirection == ShellDirection.YZ
                               select shellElement;
                var elements = new List<ShellElement>();
                foreach (ShellElement shellElement in yzshells)
                {
                    double x = shellElement.Points.Values.ToList()[0].X;
                    if (!yzElementsDic.ContainsKey(x))
                    {
                        yzElementsDic.Add(x, new List<ShellElement>() { shellElement });
                    }
                    else
                    {
                        yzElementsDic[x].Add(shellElement);
                    }
                    Plane yzP = shellElement.GetPlane();
                    if (!yzPlanes.ContainsKey(yzP.GetPlaneScript()))
                    {
                        elements.Add(shellElement);
                        yzPlanes.Add(yzP.GetPlaneScript(), elements);
                    }
                    else
                    {
                        yzPlanes[yzP.GetPlaneScript()].Add(shellElement);
                    }
                }
            }

            catch (Exception exc)
            {
                MessageBox.Show(exc.ToString());
            }

            ReOpenSap(SapFilePath);
        }
        public List<string> GetSelectedElements()
        {
            var myList = new List<string>();
            int selectedNum = 0;
            string[] selectedNames = null;
            ret = SapModel.AreaObj.GetNameList(ref selectedNum, ref selectedNames);
            for (int i = 0; i < selectedNum; i++)
            {
                bool selected = false;
                ret = SapModel.AreaObj.GetSelected(selectedNames.GetValue(i).ToString(), ref selected);
                if (selected)
                    myList.Add(selectedNames.GetValue(i).ToString());
            }
            return myList;
        }
        public bool HaveSamePlane(List<ShellElement> elements, out ShellDirection direction)
        {
            var tempList = new List<ShellDirection>();
            for (int i = 0; i < elements.Count; i++)
            {
                if (!tempList.Contains(elements[i].PlaneDirection))
                {
                    tempList.Add(elements[i].PlaneDirection);
                }
            }
            if (tempList.Count == 1)
            {
                direction = tempList[0];
                return true;
            }
            direction = tempList[0];
            return false;
        }
        public bool AreInSamePlane(List<ShellElement> elements)
        {
            var plane = elements[0].GetPlane();
            foreach (ShellElement selElement in elements)
            {
                foreach (KeyValuePair<string, Point> point in selElement.Points)
                {
                    if (!VectorOperations.IsInPlane(plane, point.Value.XYZ()))
                        return false;
                }
            }
            return true;
        }
        private bool SelElementsContainsPoint(List<ShellElement> selElements, Point point)
        {
            var points = from shellElement in selElements where shellElement.Points.ContainsValue(point) select point;
            if (points.Count() > 0)
            {
                return true;
            }
            return false;
        }
        public string GetMaterial(ShellElement element)
        {
            string propName = "";
            ret = SapModel.AreaObj.GetMaterialOverwrite(element.Name, ref propName);
            return propName;
        }
        public void GetShellProp(ShellElement element, out double thickness, out string propName)
        {
            int shellType = 0;
            string matProp = "";
            double matAng = 0;
            double bend = 0;
            int col = 0;
            string notes = "";
            string guid = "";
            double th = 0;
            string pName = "";

            ret = SapModel.AreaObj.GetProperty(element.Name, ref pName);
            ret = SapModel.PropArea.GetShell(pName, ref shellType, ref matProp, ref matAng, ref th, ref bend,
                                             ref col, ref notes, ref guid);
            propName = matProp;
            thickness = th;
        }
        public void GetStripMoments(List<ShellElement> selElements, ShellDirection direction,
                                string[] loadCases, double angle,
                                out Dictionary<string, Dictionary<string, double>> Mx,
                                out Dictionary<string, Dictionary<string, double>> My,
                                out Dictionary<string, Dictionary<string, double>> Mz,
                                out double dx, out double dy, out double dz)
        {
            Mx = new Dictionary<string, Dictionary<string, double>>();
            My = new Dictionary<string, Dictionary<string, double>>();
            Mz = new Dictionary<string, Dictionary<string, double>>();
            List<double> xVal;
            List<double> yVal;
            List<double> zVal;
            GetPointsOnAxis(selElements, out xVal, out yVal, out zVal);
            dx = xVal.Max() - xVal.Min();
            dy = yVal.Max() - yVal.Min();
            dz = zVal.Max() - zVal.Min();

            double refZ = selElements[0].Points.ToList()[0].Value.Z;
            double refX = selElements[0].Points.ToList()[0].Value.X;
            double refY = selElements[0].Points.ToList()[0].Value.Y;
            foreach (string loadCase in loadCases)
            {
                var mmx = new Dictionary<string, double>();
                var mmy = new Dictionary<string, double>();
                var mmz = new Dictionary<string, double>();

                switch (direction)
                {
                    case ShellDirection.XY:
                        for (int i = 0; i < xVal.Count; i++)
                        {
                            double my = 0;
                            var xPoints = from shellPoint in AllShellPoints
                                          where
                                              (shellPoint.Value.X == xVal[i] && shellPoint.Value.Y >= yVal.Min() &&
                                               shellPoint.Value.Y <= yVal.Max() && shellPoint.Value.Z == refZ)
                                          select shellPoint;
                            for (int j = 0; j < xPoints.ToList().Count; j++)
                            {
                                my += xPoints.ToList()[j].Value.Loads[loadCase][0] *
                                      EffectiveLength(xPoints.ToList()[j].Value.Name, "y");
                            }
                            if (!mmy.Keys.Contains(xVal[i].ToString()))
                                mmy.Add(xVal[i].ToString(), my);
                        }
                        for (int i = 0; i < yVal.Count; i++)
                        {
                            double mx = 0;
                            var yPoints = from shellPoint in AllShellPoints
                                          where
                                              (shellPoint.Value.Y == yVal[i] && shellPoint.Value.X >= xVal.Min() &&
                                               shellPoint.Value.X <= xVal.Max() && shellPoint.Value.Z == refZ)
                                          select shellPoint;
                            for (int j = 0; j < yPoints.ToList().Count; j++)
                            {
                                mx += yPoints.ToList()[j].Value.Loads[loadCase][1] *
                                      EffectiveLength(yPoints.ToList()[j].Value.Name, "x");
                            }
                            if (!mmx.Keys.Contains(yVal[i].ToString()))
                                mmx.Add(yVal[i].ToString(), mx);
                        }
                        if (angle == 90 || angle == 270)
                        {
                            var temp = mmy;
                            mmy = mmx;
                            mmx = temp;
                        }
                        Mx.Add(loadCase, mmx);
                        My.Add(loadCase, mmy);
                        Mz.Add(loadCase, mmz);

                        break;
                    case ShellDirection.XZ:
                        for (int i = 0; i < xVal.Count; i++)
                        {
                            double mz = 0;
                            var xPoints = from shellPoint in AllShellPoints
                                          where
                                              (shellPoint.Value.X == xVal[i] && shellPoint.Value.Z >= zVal.Min() &&
                                               shellPoint.Value.Z <= zVal.Max() && shellPoint.Value.Y == refY)
                                          select shellPoint;
                            for (int j = 0; j < xPoints.ToList().Count; j++)
                            {
                                mz += xPoints.ToList()[j].Value.Loads[loadCase][0] *
                                      EffectiveLength(xPoints.ToList()[j].Value.Name, "z");
                            }
                            if (!mmz.Keys.Contains(xVal[i].ToString()))
                                mmz.Add(xVal[i].ToString(), mz);
                        }
                        for (int i = 0; i < zVal.Count; i++)
                        {
                            double mx = 0;
                            var zPoints = from shellPoint in AllShellPoints
                                          where
                                              (shellPoint.Value.Z == zVal[i] && shellPoint.Value.X >= xVal.Min() &&
                                               shellPoint.Value.X <= xVal.Max() && shellPoint.Value.Y == refY)
                                          select shellPoint;
                            for (int j = 0; j < zPoints.ToList().Count; j++)
                            {
                                mx += zPoints.ToList()[j].Value.Loads[loadCase][1] *
                                      EffectiveLength(zPoints.ToList()[j].Value.Name, "x");
                            }
                            if (!mmx.Keys.Contains(zVal[i].ToString()))
                                mmx.Add(zVal[i].ToString(), mx);
                        }
                        if (angle == 90 || angle == 270)
                        {
                            var temp = mmz;
                            mmz = mmx;
                            mmx = temp;
                        }
                        Mx.Add(loadCase, mmx);
                        My.Add(loadCase, mmy);
                        Mz.Add(loadCase, mmz);
                        break;
                    case ShellDirection.YZ:
                        for (int i = 0; i < yVal.Count; i++)
                        {
                            double mz = 0;
                            var yPoints = from shellPoint in AllShellPoints
                                          where
                                              (shellPoint.Value.Y == yVal[i] && shellPoint.Value.Z >= zVal.Min() &&
                                               shellPoint.Value.Z <= zVal.Max() && shellPoint.Value.X == refX)
                                          select shellPoint;
                            for (int j = 0; j < yPoints.ToList().Count; j++)
                            {
                                mz += yPoints.ToList()[j].Value.Loads[loadCase][0] *
                                      EffectiveLength(yPoints.ToList()[j].Value.Name, "z");
                            }
                            if (!mmz.Keys.Contains(yVal[i].ToString()))
                                mmz.Add(yVal[i].ToString(), mz);
                        }
                        for (int i = 0; i < zVal.Count; i++)
                        {
                            double my = 0;
                            var zPoints = from shellPoint in AllShellPoints
                                          where
                                              (shellPoint.Value.Z == zVal[i] && shellPoint.Value.Y >= yVal.Min() &&
                                               shellPoint.Value.Y <= yVal.Max() && shellPoint.Value.X == refX)
                                          select shellPoint;
                            for (int j = 0; j < zPoints.ToList().Count; j++)
                            {
                                my += zPoints.ToList()[j].Value.Loads[loadCase][1] *
                                      EffectiveLength(zPoints.ToList()[j].Value.Name, "y");
                            }
                            if (!mmy.Keys.Contains(zVal[i].ToString()))
                                mmy.Add(zVal[i].ToString(), my);
                        }
                        if (angle == 90 || angle == 270)
                        {
                            var temp = mmz;
                            mmz = mmy;
                            mmy = temp;
                        }
                        Mx.Add(loadCase, mmx);
                        My.Add(loadCase, mmy);
                        Mz.Add(loadCase, mmz);
                        break;
                }
            }
        }
        public List<Point> GetSelPoints(List<ShellElement> selElements)
        {
            var myList = new List<Point>();
            foreach (ShellElement element in selElements)
            {
                foreach (KeyValuePair<string, Point> point in element.Points)
                {
                    if (!myList.Contains(point.Value))
                    {
                        myList.Add(point.Value);
                    }

                }
            }
            return myList;
        }

        private double EffectiveLength(string pointName, string dir)
        {
            double lx = 0;
            double ly = 0;
            double lz = 0;

            int c = PointCol[pointName].Count;

            for (int i = 0; i < PointCol[pointName].Count; i++)
            {
                lx += PointCol[pointName][i].LengthX;
            }
            for (int i = 0; i < PointCol[pointName].Count; i++)
            {
                ly += PointCol[pointName][i].LengthY;
            }
            for (int i = 0; i < PointCol[pointName].Count; i++)
            {
                lz += PointCol[pointName][i].LengthZ;
            }

            if (dir.ToLower() == "x")
            {
                if (c == 1)
                    return lx / 2;
                if (c == 2)
                    return lx / 2;
                if (c == 4)
                    return lx / 4;
                return 0;
            }
            else if (dir.ToLower() == "y")
            {
                if (c == 1)
                    return ly / 2;
                if (c == 2)
                    return ly / 2;
                if (c == 4)
                    return ly / 4;
                return 0;
            }
            else
            {
                if (c == 1)
                    return lz / 2;
                if (c == 2)
                    return lz / 2;
                if (c == 4)
                    return lz / 4;
                return 0;
            }
        }
        private double EffectiveLocalLength(string pointName, string dir)
        {
            double l11 = 0;
            double l22 = 0;

            int c = PointCol[pointName].Count;

            for (int i = 0; i < PointCol[pointName].Count; i++)
            {
                l11 += PointCol[pointName][i].LengthX;
            }

            for (int i = 0; i < PointCol[pointName].Count; i++)
            {
                l22 += PointCol[pointName][i].LengthY;
            }

            if (dir.ToLower() == "11")
            {
                if (c == 1)
                    return l11 / 2;
                if (c == 2)
                    return l11 / 2;
                if (c == 4)
                    return l11 / 4;
                return 0;
            }
            else if (dir.ToLower() == "22")
            {
                if (c == 1)
                    return l22 / 2;
                if (c == 2)
                    return l22 / 2;
                if (c == 4)
                    return l22 / 4;
                return 0;
            }
            else
            {
                return 0;
            }
        }

        private void GetPointsOnAxis(List<ShellElement> selElements, out List<double> xVal, out List<double> yVal, out List<double> zVal)
        {
            Dictionary<string, Point> selPoints = new Dictionary<string, Point>();
            foreach (ShellElement shellElement in selElements)
            {
                foreach (KeyValuePair<string, Point> point in shellElement.Points)
                {
                    if (!selPoints.ContainsKey(point.Value.Name))
                    {
                        selPoints.Add(point.Value.Name, point.Value);
                    }
                }
            }
            //List<Point> xSortedPoints = ((Dictionary<string, Point>)selPoints.OrderBy(point => point.Value.X)).Values.ToList();
            List<Point> xSortedPoints = GetMyDic(selPoints.OrderBy(point => point.Value.X)).Values.ToList();
            List<Point> ySortedPoints = GetMyDic(selPoints.OrderBy(point => point.Value.Y)).Values.ToList();
            List<Point> zSortedPoints = GetMyDic(selPoints.OrderBy(point => point.Value.Z)).Values.ToList();
            xVal = new List<double>();
            yVal = new List<double>();
            zVal = new List<double>();
            for (int i = 0; i < xSortedPoints.Count; i++)
            {
                if (!xVal.Contains(xSortedPoints[i].X))
                    xVal.Add(xSortedPoints[i].X);
                if (!yVal.Contains(ySortedPoints[i].Y))
                    yVal.Add(ySortedPoints[i].Y);
                if (!zVal.Contains(zSortedPoints[i].Z))
                    zVal.Add(zSortedPoints[i].Z);
            }
        }
        internal static Dictionary<string, Point> GetMyDic(IEnumerable<KeyValuePair<string, Point>> iEnum)
        {
            var myDic = new Dictionary<string, Point>();
            var en = iEnum.GetEnumerator();
            while (en.MoveNext())
            {
                myDic.Add(en.Current.Key, en.Current.Value);
            }
            return myDic;
        }
    }
}
