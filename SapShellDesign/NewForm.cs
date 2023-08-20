using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Tao.OpenGl;

namespace SapShellDesign
{
    public partial class NewForm : Form
    {
        public NewForm()
        {
            InitializeComponent();
        }

        private SapShell sapShell;

        private double dimX;
        private double dimY;
        private double dimZ;
        private double th;
        private string propName;
        private bool SelectionDone;
        Dictionary<string, Dictionary<string, double>> Mx;
        Dictionary<string, Dictionary<string, double>> My;
        Dictionary<string, Dictionary<string, double>> Mz;
        Dictionary<string, Dictionary<string, Dictionary<string, double>>> MxStripReinArea;
        Dictionary<string, Dictionary<string, Dictionary<string, double>>> MyStripReinArea;
        Dictionary<string, Dictionary<string, Dictionary<string, double>>> MzStripReinArea;
        Dictionary<string, Dictionary<string, List<ContourPos>>> M11PointReinArea;
        Dictionary<string, Dictionary<string, List<ContourPos>>> M22PointReinArea;
        Dictionary<string, ShellElement> allElements;
        Dictionary<double, List<ShellElement>> xyElements;
        Dictionary<double, List<ShellElement>> xzElements;
        Dictionary<double, List<ShellElement>> yzElements;
        private Dictionary<string, List<ShellElement>> xyPlanes;
        private Dictionary<string, List<ShellElement>> xzPlanes;
        private Dictionary<string, List<ShellElement>> yzPlanes;
        private bool ReinBarIsVertcal;
        private double CurrentBe;

        private Dictionary<string, Dictionary<string, Dictionary<string, double>>> BuildStripReinAreaDictionaries(string[] desingMethod, Dictionary<string, Dictionary<string, double>> m, string axis)
        {
            var myDic = new Dictionary<string, Dictionary<string, Dictionary<string, double>>>();

            double be;
            bool vert = false;
            switch (axis)
            {
                case "x":
                    {
                        be = dimX;
                        if (planeDirection == ShellDirection.XZ)
                        {
                            vert = true;
                        }
                        break;
                    }
                case "y":
                    {
                        be = dimY;
                        if (planeDirection == ShellDirection.YZ)
                        {
                            vert = true;
                        }
                        break;
                    }
                case "z":
                    {
                        be = dimZ;
                        break;
                    }
                default:
                    be = 0;
                    break;
            }
            for (int i = 0; i < desingMethod.Length; i++)
            {
                var reinAreaDic = new Dictionary<string, Dictionary<string, double>>();
                foreach (KeyValuePair<string, Dictionary<string, double>> pair in m)
                {
                    var reinValueDic = new Dictionary<string, double>();
                    List<double> moments = pair.Value.Values.ToList();
                    var reinAreaValues = new List<double>();
                    foreach (KeyValuePair<string, double> valuePair in pair.Value)
                    {
                        reinValueDic.Add(valuePair.Key,
                                         ConvertToReinforceArea(desingMethod.GetValue(i).ToString(), valuePair.Value,
                                                                Convert.ToDouble(txtFc.Text),
                                                                Convert.ToDouble(txtFs.Text),
                                                                Convert.ToDouble(txtFy.Text), be,
                                                                th, cmbTempReinbar.SelectedIndex, vert));
                    }

                    reinAreaDic.Add(pair.Key, reinValueDic);
                }
                myDic.Add(desingMethod.GetValue(i).ToString(), reinAreaDic);
            }
            return myDic;
        }
        private Dictionary<string, Dictionary<string, List<ContourPos>>> BuildPointReinAreaDictionaries(string desingMethod, List<Point> selPoints, string localM)
        {
            var myDic = new Dictionary<string, Dictionary<string, List<ContourPos>>>();
            int mInd = 0;
            if (localM.ToLower() == "m11")
                mInd = 0;
            else
                mInd = 1;
            double be = 1;
            var reinAreaDic = new Dictionary<string, List<ContourPos>>();
            foreach (Point selPoint in selPoints)
            {
                foreach (KeyValuePair<string, double[]> load in selPoint.Loads)
                {
                    if (!reinAreaDic.ContainsKey(load.Key))
                    {
                        var list = new List<ContourPos>();
                        list.Add(new ContourPos()
                        {
                            X = selPoint.X,
                            Y = selPoint.Y,
                            Z = selPoint.Z,

                            Value =
                                             ConvertToReinforceArea(desingMethod, load.Value[mInd],
                                                                    Convert.ToDouble(txtFc.Text),
                                                                    Convert.ToDouble(txtFs.Text),
                                                                    Convert.ToDouble(txtFy.Text), be, th, cmbTempReinbar.SelectedIndex, false)
                        });
                        reinAreaDic.Add(load.Key, list);
                    }
                    else
                    {
                        var newPos = new ContourPos()
                        {
                            X = selPoint.X,
                            Y = selPoint.Y,
                            Z = selPoint.Z,

                            Value =
                                                 ConvertToReinforceArea(desingMethod, load.Value[mInd],
                                                                        Convert.ToDouble(txtFc.Text),
                                                                        Convert.ToDouble(txtFs.Text),
                                                                        Convert.ToDouble(txtFy.Text), be, th, cmbTempReinbar.SelectedIndex, false)
                        };
                        reinAreaDic[load.Key].Add(newPos);
                    }
                }
            }
            myDic.Add(desingMethod, reinAreaDic);
            return myDic;
        }

        private string[] ComboNames;

        private void NewForm_Load(object sender, EventArgs e)
        {
            // simpleOpenGlControl1.InitializeContexts();
            this.Text = "SAP2000 Shell Designer (Ver." + Application.ProductVersion + ")";

        }
        void InitializeComboBoxes()
        {

            cmbComboNames.Items.Clear();
            chkCombosListDesign.Items.Clear();
            chkComboListCrack.Items.Clear();

            for (int i = 0; i < ComboNames.Length; i++)
            {
                cmbComboNames.Items.Add(ComboNames[i]);
                chkCombosListDesign.Items.Add(ComboNames[i]);
                chkComboListCrack.Items.Add(ComboNames[i]);
            }

            if (cmbComboNames.Items.Count > 0)
                cmbComboNames.SelectedIndex = 0;
            if (cmbStripMomentType.Items.Count > 0)
                cmbStripMomentType.SelectedIndex = 0;
            if (cmbStripOutputType.Items.Count > 0)
                cmbStripOutputType.SelectedIndex = 0;
            if (cmbReinbars.Items.Count > 0)
                cmbReinbars.SelectedIndex = 0;
            if (cmbDesignMethod.Items.Count > 0)
                cmbDesignMethod.SelectedIndex = 0;
            cmbTempReinbar.SelectedIndex = 0;
        }

        private void simpleOpenGlControl1_Paint(object sender, PaintEventArgs e)
        {
            PaintIt();
        }

        private void simpleOpenGlControl1_Resize(object sender, EventArgs e)
        {
            //  simpleOpenGlControl1.Invalidate();
            PaintIt();
        }
        private void PaintIt()
        {
            Gl.glClearColor(0.0f, 0.0f, 0.0f, 0.0f);
            Gl.glClear(Gl.GL_COLOR_BUFFER_BIT | Gl.GL_DEPTH_BUFFER_BIT);
            Gl.glOrtho(-.5, 1, -.5, 1, -1, 1);

            //Gl.glEnable(Gl.GL_BLEND);
            //Gl.glBlendFunc(Gl.GL_BLEND_SRC_ALPHA, Gl.GL_ONE_MINUS_SRC_ALPHA);
            //Gl.glShadeModel(Gl.GL_FLAT);


            Gl.glBegin(Gl.GL_POLYGON);
            {
                Gl.glColor3f(0f, 0f, 0f); Gl.glVertex2f(0.10f, 0.10f);
                Gl.glColor3f(0f, 0f, 0f); Gl.glVertex2f(0.70f, 0.10f);
                Gl.glColor3f(1f, 1f, 1f); Gl.glVertex2f(0.70f, 0.3f);
                Gl.glColor3f(1f, 1f, 1f); Gl.glVertex2f(0.10f, 0.30f);
            }
            Gl.glEnd();
            Gl.glBegin(Gl.GL_POLYGON);
            {
                Gl.glColor3f(1f, 1f, 1f); Gl.glVertex2f(0.10f, 0.30f);
                Gl.glColor3f(1f, 1f, 1f); Gl.glVertex2f(0.70f, 0.30f);
                Gl.glColor3f(0f, 0f, 0f); Gl.glVertex2f(0.70f, 0.6f);
                Gl.glColor3f(0f, 0f, 0f); Gl.glVertex2f(0.10f, 0.60f);
            }
            Gl.glEnd();
            Gl.glFlush();
        }

        public void btnOpenFile_Click(object sender, EventArgs e)
        {
            var fd = new OpenFileDialog();
            fd.Filter = "Sap2000 Files|*.sdb";
            System.Windows.Forms.DialogResult dialogResult = fd.ShowDialog();
            if (dialogResult != System.Windows.Forms.DialogResult.Cancel)
            {
                sapShell = new SapShell(fd.FileName);
                if (sapShell != null)
                {
                    sapShell.GetShells(out allElements, out xyPlanes, out xzPlanes, out yzPlanes, out ComboNames);
                    groupBox4.Enabled = true;
                    InitializeComboBoxes();
                }

                //sapShell = null;
            }
        }

        private List<string> selList;
        public double ConverToDouble(string s)
        {
            return Convert.ToDouble(s);
        }
        public double RoundAll(double d)
        {
            return Math.Round(d, 2);
        }
        public double ConvertToReinforceArea(string desMethod, double m, double fc, double fs, double fy, double be, double th, int tempAreaInd, bool vert)
        {
            if (desMethod == "Allowable Stress")
            {
                double asReq = m / (fs * 7 / 8.0 * (th - .06)) * 10000;
                double asMin = 14.0 / (fy / 10.0) * be * (th - .06) * 10000 * Math.Sign(m);
                if (Math.Abs(asMin) < Math.Abs(asReq))
                {
                    return asReq;
                }
                if (Math.Abs(asMin) > 1.33 * Math.Abs(asReq))
                {
                    return 1.33 * asReq;
                }
                return asMin;
            }
            if (desMethod == "ACI")
            {
                return AciRectBeamDesign(m, be, th, Convert.ToDouble(txtCoverPos.Text.Trim()), fc, fy, tempAreaInd, vert);
            }
            return 0;
        }
        ShellDirection planeDirection;
        private void btnRefreshStripSelection_Click(object sender, EventArgs e)
        {
            SelectionDone = false;
            selList = sapShell.GetSelectedElements();
            if (selList != null)
            {
                FillObjects();
            }
            else
                MessageBox.Show("No Elements Selected!");
        }

        private List<ShellElement> selElements;
        private void FillObjects()
        {

            selElements = new List<ShellElement>();
            foreach (string s in selList)
            {
                selElements.Add(allElements[s]);
            }
            if (sapShell.HaveSamePlane(selElements, out planeDirection))
            {
                groupBox1.Enabled = true;
                printChartToolStripMenuItem.Enabled = true;
                toolStripMenuItem2.Enabled = true;
                //InitializeComboBoxes();
                if (planeDirection == ShellDirection.XY || planeDirection == ShellDirection.XZ || planeDirection == ShellDirection.YZ)
                {
                    sapShell.GetStripMoments(selElements, planeDirection, ComboNames, 0, out Mx, out My, out Mz, out dimX, out dimY, out dimZ);

                    sapShell.GetShellProp(selElements[0], out th, out propName);
                    string stripDimStr = "";
                    if (planeDirection == ShellDirection.XY)
                    {
                        stripDimStr = (Math.Round(dimX, 2) + " X " + Math.Round(dimY, 2));
                    }
                    else if (planeDirection == ShellDirection.XZ)
                    {
                        stripDimStr = (Math.Round(dimX, 2) + " X " + Math.Round(dimZ, 2));
                    }
                    if (planeDirection == ShellDirection.YZ)
                    {
                        stripDimStr = (Math.Round(dimY, 2) + " X " + Math.Round(dimZ, 2));
                    }

                    MxStripReinArea = BuildStripReinAreaDictionaries(new[] { "Allowable Stress", "ACI" }, Mx, "x");
                    MyStripReinArea = BuildStripReinAreaDictionaries(new[] { "Allowable Stress", "ACI" }, My, "y");
                    MzStripReinArea = BuildStripReinAreaDictionaries(new[] { "Allowable Stress", "ACI" }, Mz, "z");

                    UpdateStatus(planeDirection.ToString(), selElements.Count,
                                 stripDimStr, propName, th.ToString());
                    SelectionDone = true;
                    UpdateChart();
                }
            }

        }
        private double CurrentStripWidth;
        private double CurrentStripHeight;
        private double CurrentStripLength;

        void UpdateStatus(string direction, int numOfElements, string stripDim, string stripMaterial, string stripThickness)
        {
            txtStripPlane.Text = direction;
            txtStripElementsNumber.Text = numOfElements.ToString();
            txtStripDimension.Text = stripDim;
            txtStripMaterial.Text = stripMaterial;
            txtStripThickness.Text = stripThickness;
        }

        void UpdateChart()
        {
            if (SelectionDone)
            {
                chart1.Series[0].Points.Clear();
                chart1.Series[1].Points.Clear();

                if (cmbStripMomentType.SelectedIndex == 0)
                {
                    chart1.Titles[0].Text = "X-X";
                    switch (cmbStripOutputType.SelectedIndex)
                    {
                        case 0:
                            chart1.ChartAreas[0].AxisY.Title = "Moment (t.m)";
                            chart1.Series[0].Points.DataBindXY(Mx[cmbComboNames.Text].Keys.ToList().ConvertAll(ConverToDouble).ToArray(),
                                                               Mx[cmbComboNames.Text].Values.ToArray());
                            return;
                        case 1:
                            chart1.ChartAreas[0].AxisY.Title = "Reinforce Area (cm2)";
                            chart1.Series[0].Points.DataBindXY(MxStripReinArea[cmbDesignMethod.Text][cmbComboNames.Text].Keys.ToList().ConvertAll(ConverToDouble).ToArray(),
                                                               MxStripReinArea[cmbDesignMethod.Text][cmbComboNames.Text].Values.ToArray());
                            return;
                        case 2:
                            chart1.ChartAreas[0].AxisY.Title = "Rebar No.";
                            chart1.Series[0].Points.DataBindXY(MxStripReinArea[cmbDesignMethod.Text][cmbComboNames.Text].Keys.ToList().ConvertAll(ConverToDouble).ToArray(),
                                                               MxStripReinArea[cmbDesignMethod.Text][cmbComboNames.Text].Values.ToList().ConvertAll(ConvertToRebarNumber).ToArray());
                            return;
                    }
                }
                if (cmbStripMomentType.SelectedIndex == 1)
                {
                    chart1.Titles[0].Text = "Y-Y";
                    switch (cmbStripOutputType.SelectedIndex)
                    {
                        case 0:
                            chart1.ChartAreas[0].AxisY.Title = "Moment (t.m)";
                            chart1.Series[0].Points.DataBindXY(My[cmbComboNames.Text].Keys.ToList().ConvertAll(ConverToDouble).ToArray(),
                                                               My[cmbComboNames.Text].Values.ToArray());
                            return;
                        case 1:
                            chart1.ChartAreas[0].AxisY.Title = "Reinforce Area (cm2)";
                            chart1.Series[0].Points.DataBindXY(MyStripReinArea[cmbDesignMethod.Text][cmbComboNames.Text].Keys.ToList().ConvertAll(ConverToDouble).ToArray(),
                                                               MyStripReinArea[cmbDesignMethod.Text][cmbComboNames.Text].Values.ToArray());
                            return;
                        case 2:
                            chart1.ChartAreas[0].AxisY.Title = "Rebar No.";
                            chart1.Series[0].Points.DataBindXY(MyStripReinArea[cmbDesignMethod.Text][cmbComboNames.Text].Keys.ToList().ConvertAll(ConverToDouble).ToArray(),
                                                               MyStripReinArea[cmbDesignMethod.Text][cmbComboNames.Text].Values.ToList().ConvertAll(ConvertToRebarNumber).ToArray());
                            return;
                    }
                }
                if (cmbStripMomentType.SelectedIndex == 2)
                {
                    chart1.Titles[0].Text = "Z-Z";
                    switch (cmbStripOutputType.SelectedIndex)
                    {
                        case 0:
                            chart1.ChartAreas[0].AxisY.Title = "Moment (t.m)";
                            chart1.Series[0].Points.DataBindXY(Mz[cmbComboNames.Text].Keys.ToList().ConvertAll(ConverToDouble).ToArray(),
                                                               Mz[cmbComboNames.Text].Values.ToArray());
                            return;
                        case 1:
                            chart1.ChartAreas[0].AxisY.Title = "Reinforce Area (cm2)";
                            chart1.Series[0].Points.DataBindXY(MzStripReinArea[cmbDesignMethod.Text][cmbComboNames.Text].Keys.ToList().ConvertAll(ConverToDouble).ToArray(),
                                                               MzStripReinArea[cmbDesignMethod.Text][cmbComboNames.Text].Values.ToArray());
                            return;
                        case 2:
                            chart1.ChartAreas[0].AxisY.Title = "Rebar No.";
                            chart1.Series[0].Points.DataBindXY(MzStripReinArea[cmbDesignMethod.Text][cmbComboNames.Text].Keys.ToList().ConvertAll(ConverToDouble).ToArray(),
                                                               MzStripReinArea[cmbDesignMethod.Text][cmbComboNames.Text].Values.ToList().ConvertAll(ConvertToRebarNumber).ToArray());
                            return;
                    }
                }
            }
        }

        private Dictionary<string, double> envelopeMaxServiceMx;
        private Dictionary<string, double> envelopeMaxServiceMy;
        private Dictionary<string, double> envelopeMaxServiceMz;

        private Dictionary<string, double> envelopeMinServiceMx;
        private Dictionary<string, double> envelopeMinServiceMy;
        private Dictionary<string, double> envelopeMinServiceMz;


        private Dictionary<string, double> envelopeMaxMx;
        private Dictionary<string, double> envelopeMaxMy;
        private Dictionary<string, double> envelopeMaxMz;

        private Dictionary<string, double> envelopeMinMx;
        private Dictionary<string, double> envelopeMinMy;
        private Dictionary<string, double> envelopeMinMz;

        private Dictionary<string, double> envelopeMaxRx;
        private Dictionary<string, double> envelopeMaxRy;
        private Dictionary<string, double> envelopeMaxRz;

        private Dictionary<string, double> envelopeMinRx;
        private Dictionary<string, double> envelopeMinRy;
        private Dictionary<string, double> envelopeMinRz;

        private EnvelopeType CurrentEnvelope;
        private enum EnvelopeType
        {
            envelopeMx,
            envelopeMy,
            envelopeMz,
        }
        void UpdateChartForEnvelope()
        {
            if (SelectionDone)
            {
                envelopeMaxServiceMx = GetEnvelopeServiceMoments(Mx, true);
                envelopeMaxServiceMy = GetEnvelopeServiceMoments(My, true);
                envelopeMaxServiceMz = GetEnvelopeServiceMoments(Mz, true);

                envelopeMinServiceMx = GetEnvelopeServiceMoments(Mx, false);
                envelopeMinServiceMy = GetEnvelopeServiceMoments(My, false);
                envelopeMinServiceMz = GetEnvelopeServiceMoments(Mz, false);


                envelopeMaxMx = GetEnvelopeMoments(Mx, true);
                envelopeMaxMy = GetEnvelopeMoments(My, true);
                envelopeMaxMz = GetEnvelopeMoments(Mz, true);

                envelopeMinMx = GetEnvelopeMoments(Mx, false);
                envelopeMinMy = GetEnvelopeMoments(My, false);
                envelopeMinMz = GetEnvelopeMoments(Mz, false);

                envelopeMaxRx = GetEnvelopeReinBars(MxStripReinArea, true);
                envelopeMaxRy = GetEnvelopeReinBars(MyStripReinArea, true);
                envelopeMaxRz = GetEnvelopeReinBars(MzStripReinArea, true);

                envelopeMinRx = GetEnvelopeReinBars(MxStripReinArea, false);
                envelopeMinRy = GetEnvelopeReinBars(MyStripReinArea, false);
                envelopeMinRz = GetEnvelopeReinBars(MzStripReinArea, false);

                chart1.Series[0].Points.Clear();
                chart1.Series[1].Points.Clear();
                if (cmbStripMomentType.SelectedIndex == 0)
                {
                    chart1.Titles[0].Text = "Results Around Axis X-X";
                    switch (cmbStripOutputType.SelectedIndex)
                    {
                        case 0:
                            chart1.ChartAreas[0].AxisY.Title = "Envelope Moment (t.m)";
                            chart1.Series[0].Points.DataBindXY(
                                envelopeMaxMx.Keys.ToList().ConvertAll(ConverToDouble).ToArray(),
                                envelopeMaxMx.Values.ToList().ConvertAll(RoundAll).ToArray());
                            chart1.Series[1].Points.DataBindXY(
                                envelopeMinMx.Keys.ToList().ConvertAll(ConverToDouble).ToArray(),
                                envelopeMinMx.Values.ToList().ConvertAll(RoundAll).ToArray());
                            CurrentEnvelope = EnvelopeType.envelopeMx;
                            return;

                        case 1:
                            chart1.ChartAreas[0].AxisY.Title = "Reinforce Area (cm2)";
                            chart1.Series[0].Points.DataBindXY(
                                envelopeMaxRx.Keys.ToList().ConvertAll(ConverToDouble).ToArray(),
                                envelopeMaxRx.Values.ToList().ConvertAll(RoundAll).ToArray());
                            chart1.Series[1].Points.DataBindXY(
                                envelopeMinRx.Keys.ToList().ConvertAll(ConverToDouble).ToArray(),
                                envelopeMinRx.Values.ToList().ConvertAll(RoundAll).ToArray());
                            CurrentEnvelope = EnvelopeType.envelopeMx;
                            return;

                        case 2:
                            chart1.ChartAreas[0].AxisY.Title = "Rebar No.";
                            return;
                    }
                }
                if (cmbStripMomentType.SelectedIndex == 1)
                {
                    chart1.Titles[0].Text = "Results Arounds Y-Y";
                    switch (cmbStripOutputType.SelectedIndex)
                    {
                        case 0:
                            chart1.ChartAreas[0].AxisY.Title = "Moment (t.m)";
                            chart1.Series[0].Points.DataBindXY(
                                envelopeMaxMy.Keys.ToList().ConvertAll(ConverToDouble).ToArray(),
                                envelopeMaxMy.Values.ToList().ConvertAll(RoundAll).ToArray());
                            chart1.Series[1].Points.DataBindXY(
                                envelopeMinMy.Keys.ToList().ConvertAll(ConverToDouble).ToArray(),
                                envelopeMinMy.Values.ToList().ConvertAll(RoundAll).ToArray());
                            CurrentEnvelope = EnvelopeType.envelopeMy;
                            return;
                        case 1:
                            chart1.ChartAreas[0].AxisY.Title = "Reinforce Area (cm2)";
                            chart1.Series[0].Points.DataBindXY(
                                envelopeMaxRy.Keys.ToList().ConvertAll(ConverToDouble).ToArray(),
                                envelopeMaxRy.Values.ToList().ConvertAll(RoundAll).ToArray());
                            chart1.Series[1].Points.DataBindXY(
                                envelopeMinRy.Keys.ToList().ConvertAll(ConverToDouble).ToArray(),
                                envelopeMinRy.Values.ToList().ConvertAll(RoundAll).ToArray());
                            CurrentEnvelope = EnvelopeType.envelopeMy;
                            return;
                        case 2:
                            chart1.ChartAreas[0].AxisY.Title = "Rebar No.";
                            return;
                    }
                }
                if (cmbStripMomentType.SelectedIndex == 2)
                {
                    chart1.Titles[0].Text = "Results Around Z-Z";
                    switch (cmbStripOutputType.SelectedIndex)
                    {
                        case 0:
                            chart1.ChartAreas[0].AxisY.Title = "Moment (t.m)";
                            chart1.Series[0].Points.DataBindXY(
                                envelopeMaxMz.Keys.ToList().ConvertAll(ConverToDouble).ToArray(),
                                envelopeMaxMz.Values.ToList().ConvertAll(RoundAll).ToArray());
                            chart1.Series[1].Points.DataBindXY(
                                envelopeMinMz.Keys.ToList().ConvertAll(ConverToDouble).ToArray(),
                                envelopeMinMz.Values.ToList().ConvertAll(RoundAll).ToArray());
                            CurrentEnvelope = EnvelopeType.envelopeMz;
                            return;
                        case 1:
                            chart1.ChartAreas[0].AxisY.Title = "Reinforce Area (cm2)";
                            chart1.Series[0].Points.DataBindXY(
                                envelopeMaxRz.Keys.ToList().ConvertAll(ConverToDouble).ToArray(),
                                envelopeMaxRz.Values.ToList().ConvertAll(RoundAll).ToArray());
                            chart1.Series[1].Points.DataBindXY(
                                envelopeMinRz.Keys.ToList().ConvertAll(ConverToDouble).ToArray(),
                                envelopeMinRz.Values.ToList().ConvertAll(RoundAll).ToArray());
                            CurrentEnvelope = EnvelopeType.envelopeMz;
                            return;
                        case 2:
                            chart1.ChartAreas[0].AxisY.Title = "Rebar No.";
                            return;
                    }
                }
            }
        }

        private double ConvertToRebarNumber(double input)
        {
            double barDiam =
                Math.Abs(
                    Convert.ToDouble(
                        cmbReinbars.Text.Split(new string[] { "F" }, StringSplitOptions.RemoveEmptyEntries)[0]) / 10);
            double area = Math.PI * barDiam * barDiam / 4;

            return (int)(input / area) + 1;
        }


        private void cmbStripOutputType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbStripOutputType.SelectedIndex == 1 || cmbStripOutputType.SelectedIndex == 2)
            {
                pnlDesign.Enabled = true;
            }
            else
            {
                pnlDesign.Enabled = false;
            }
        }


        private void ReDesign()
        {
            MxStripReinArea = BuildStripReinAreaDictionaries(new[] { "Allowable Stress", "ACI" }, Mx, "x");
            MyStripReinArea = BuildStripReinAreaDictionaries(new[] { "Allowable Stress", "ACI" }, My, "y");
            MzStripReinArea = BuildStripReinAreaDictionaries(new[] { "Allowable Stress", "ACI" }, Mz, "z");
        }

        private void tsbShowHideValues_Click(object sender, EventArgs e)
        {
            chart1.Series[0].IsValueShownAsLabel = !chart1.Series[0].IsValueShownAsLabel;
            chart1.Series[1].IsValueShownAsLabel = !chart1.Series[1].IsValueShownAsLabel;
        }
        public void tsbPreview_Click(object sender, EventArgs e)
        {
            chart1.Printing.PrintPreview();
        }
        public void tsbPrint_Click(object sender, EventArgs e)
        {
            printDocument1.OriginAtMargins = true;
            printDocument1.Print();
            //chart1.Printing.Print(false);
        }
        private void openSAP2000FileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            btnOpenFile_Click(sender, e);
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            tsbPreview_Click(sender, e);
        }

        private void printChartToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tsbPrint_Click(sender, e);
        }

        private void contentsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process.Start(Application.StartupPath + @"/Help Files/SAP2000 Shell Designer Guide.pdf");
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new AboutBox1().ShowDialog();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SelectionDone = false;
            var selList = sapShell.GetSelectedElements();
            var selElements = new List<ShellElement>();
            if (selList != null)
            {
                foreach (string s in selList)
                {
                    selElements.Add(allElements[s]);
                }
                List<Point> selPoints = sapShell.GetSelPoints(selElements);

                if (!sapShell.AreInSamePlane(selElements))
                {
                    groupBox1.Enabled = true;
                    printChartToolStripMenuItem.Enabled = true;
                    toolStripMenuItem2.Enabled = true;
                    InitializeComboBoxes();

                    sapShell.GetShellProp(selElements[0], out th, out propName);
                    string stripDimStr = "";

                    M11PointReinArea = BuildPointReinAreaDictionaries("Allowable Stress", selPoints, "m11");
                    M22PointReinArea = BuildPointReinAreaDictionaries("Allowable Stress", selPoints, "m22");

                    SelectionDone = true;
                    // UpdateChart();

                }
            }
            else
                MessageBox.Show("No Elements Selected!");
        }

        private void chartToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Form2 form2 = new Form2(xyPlanes, xzPlanes, yzPlanes, sapShell);
            form2.ShowDialog();
        }
        private Dictionary<string, double> GetEnvelopeServiceMoments(Dictionary<string, Dictionary<string, double>> sourceM, bool max)
        {
            var envelope = new Dictionary<string, double>();
            for (int i = 0; i < sourceM.Values.ToList()[0].Values.Count; i++)
            {
                envelope.Add(sourceM.Values.ToList()[0].Keys.ToList()[i], 0);
            }

            foreach (KeyValuePair<string, Dictionary<string, double>> m in sourceM)
            {
                for (int i = 0; i < chkComboListCrack.CheckedItems.Count; i++)
                {
                    if (m.Key == chkComboListCrack.GetItemText(chkComboListCrack.CheckedItems[i]))
                    {
                        foreach (KeyValuePair<string, double> d in m.Value)
                        {
                            if (max)
                            {
                                if (envelope[d.Key] < d.Value)
                                {
                                    envelope[d.Key] = d.Value;
                                }
                            }
                            else
                            {
                                if (envelope[d.Key] > d.Value)
                                {
                                    envelope[d.Key] = d.Value;
                                }
                            }
                        }
                    }
                }
            }
            return envelope;
        }
        private Dictionary<string, double> GetEnvelopeMoments(Dictionary<string, Dictionary<string, double>> sourceM, bool max)
        {
            var envelope = new Dictionary<string, double>();
            for (int i = 0; i < sourceM.Values.ToList()[0].Values.Count; i++)
            {
                envelope.Add(sourceM.Values.ToList()[0].Keys.ToList()[i], 0);
            }

            foreach (KeyValuePair<string, Dictionary<string, double>> m in sourceM)
            {
                for (int i = 0; i < chkCombosListDesign.CheckedItems.Count; i++)
                {
                    if (m.Key == chkCombosListDesign.GetItemText(chkCombosListDesign.CheckedItems[i]))
                    {
                        foreach (KeyValuePair<string, double> d in m.Value)
                        {
                            if (max)
                            {
                                if (envelope[d.Key] < d.Value)
                                {
                                    envelope[d.Key] = d.Value;
                                }
                            }
                            else
                            {
                                if (envelope[d.Key] > d.Value)
                                {
                                    envelope[d.Key] = d.Value;
                                }
                            }
                        }
                    }
                }
            }
            return envelope;
        }
        private double GetTempratureArea(double b, double h, bool vert)
        {
            LimitationForTempRebar = false;
            double tempAs = 0;
            switch (cmbTempReinbar.SelectedIndex)
            {
                case 0:// General 0.0028
                    if (h > 60)
                    {
                        tempAs = 0.20 * b;
                        LimitationForTempRebar = true;
                    }
                    else
                        tempAs = 0.0028 * b * h / 2;
                    break;
                case 1: // Walls (Vertical 0.0015 Horz. 0.0025)
                    if (vert)
                    {
                        if (h > 60)
                        {
                            tempAs = 0.2 * b;
                            LimitationForTempRebar = true;
                        }
                        else
                            tempAs = 0.0015 * b * h / 2;
                    }
                    else
                    {
                        if (h > 60)
                        {
                            tempAs = 0.2 * b;
                            LimitationForTempRebar = true;
                        }
                        else
                            tempAs = 0.0025 * b * h / 2;
                    }
                    break;
                case 2:// Watertight 0.003
                    if (h > 60)
                    {
                        tempAs = .003 * b * 30;
                        LimitationForTempRebar = true;
                    }
                    else
                        tempAs = 0.003 * b * h / 2;
                    break;
                case 3:// Watertight 0.004
                    if (h > 60)
                    {
                        tempAs = .004 * b * 30;
                        LimitationForTempRebar = true;
                    }
                    else
                        tempAs = 0.004 * b * h / 2;
                    break;
                case 4:// Watertight 0.005
                    if (h > 60)
                    {
                        tempAs = .005 * b * 30;
                        LimitationForTempRebar = true;
                    }
                    else
                        tempAs = 0.005 * b * h / 2;
                    break;
                case 5:// Watertight 0.006
                    if (h > 60)
                    {
                        tempAs = .006 * b * 30;
                        LimitationForTempRebar = true;
                    }
                    else
                        tempAs = 0.006 * b * h / 2;
                    break;
            }
            return tempAs;
        }
        private Dictionary<string, double> GetEnvelopeReinBars(Dictionary<string, Dictionary<string, Dictionary<string, double>>> sourceR, bool max)
        {
            bool vert = false;
            if (cmbStripMomentType.SelectedIndex == 0)
            {
                CurrentStripWidth = dimX;
                if (planeDirection == ShellDirection.XZ)
                {
                    vert = true;
                }
            }
            if (cmbStripMomentType.SelectedIndex == 1)
            {
                CurrentStripWidth = dimY;
                if (planeDirection == ShellDirection.YZ)
                {
                    vert = true;
                }
            }
            if (cmbStripMomentType.SelectedIndex == 2)
            {
                CurrentStripWidth = dimZ;
                vert = false;
            }
            double tempAs = 0;
            double h = th * 100;
            double b = CurrentStripWidth * 100;
            tempAs = GetTempratureArea(b, h, vert);
            var envelope = new Dictionary<string, double>();
            int sgn = -1;
            if (max) sgn = 1;

            for (int i = 0; i < sourceR[cmbDesignMethod.Text].Values.ToList()[0].Values.Count; i++)
            {
                envelope.Add(sourceR[cmbDesignMethod.Text].Values.ToList()[0].Keys.ToList()[i], tempAs * sgn);
            }

            foreach (KeyValuePair<string, Dictionary<string, double>> r in sourceR[cmbDesignMethod.Text])
            {
                for (int i = 0; i < chkCombosListDesign.CheckedItems.Count; i++)
                {
                    if (r.Key == chkCombosListDesign.GetItemText(chkCombosListDesign.CheckedItems[i]))
                    {
                        foreach (KeyValuePair<string, double> d in r.Value)
                        {
                            if (max)
                            {
                                if (envelope[d.Key] < d.Value)
                                {
                                    envelope[d.Key] = d.Value;
                                }
                            }
                            else
                            {
                                if (envelope[d.Key] > d.Value)
                                {
                                    envelope[d.Key] = d.Value;
                                }
                            }
                        }
                    }
                }
            }
            return envelope;
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            ReDesign();
            UpdateChartForEnvelope();
        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            double be = dimX;
            double MaxAsPos = 0;
            double MaxAsNeg = 0;
            double MaxMomentServicePos = 0;
            double MaxMomentServiceNeg = 0;

            if (CurrentEnvelope == EnvelopeType.envelopeMx)
            {
                be = dimX;
                if (envelopeMaxServiceMx != null)
                    MaxMomentServicePos = Math.Round(envelopeMaxServiceMx.Values.ToList().Max(), 2);
                else
                    MaxMomentServicePos = 0;
                if (envelopeMinServiceMx != null)
                    MaxMomentServiceNeg = Math.Abs(Math.Round(envelopeMinServiceMx.Values.ToList().Min(), 2));
                else
                    MaxMomentServiceNeg = 0;

                if (envelopeMaxRx != null)
                    MaxAsPos = Math.Round(envelopeMaxRx.Values.ToList().Max(), 2);
                else
                    MaxAsPos = 0;
                if (envelopeMinRx != null)
                    MaxAsNeg = Math.Abs(Math.Round(envelopeMinRx.Values.ToList().Min(), 2));
                else
                    MaxAsNeg = 0;
            }
            if (CurrentEnvelope == EnvelopeType.envelopeMy)
            {
                be = dimY;
                if (envelopeMaxServiceMy != null)
                    MaxMomentServicePos = Math.Round(envelopeMaxServiceMy.Values.ToList().Max(), 2);
                else
                    MaxMomentServicePos = 0;
                if (envelopeMinServiceMy != null)
                    MaxMomentServiceNeg = Math.Abs(Math.Round(envelopeMinServiceMy.Values.ToList().Min(), 2));
                else
                    MaxMomentServiceNeg = 0;

                if (envelopeMaxRy != null)
                    MaxAsPos = Math.Round(envelopeMaxRy.Values.ToList().Max(), 2);
                else
                    MaxAsPos = 0;
                if (envelopeMinRy != null)
                    MaxAsNeg = Math.Abs(Math.Round(envelopeMinRy.Values.ToList().Min(), 2));
                else
                    MaxAsNeg = 0;
            }
            if (CurrentEnvelope == EnvelopeType.envelopeMz)
            {
                be = dimZ;
                if (envelopeMaxServiceMz != null)
                    MaxMomentServicePos = Math.Round(envelopeMaxServiceMz.Values.ToList().Max(), 2);
                else
                    MaxMomentServicePos = 0;
                if (envelopeMinServiceMz != null)
                    MaxMomentServiceNeg = Math.Abs(Math.Round(envelopeMinServiceMz.Values.ToList().Min(), 2));
                else
                    MaxMomentServiceNeg = 0;

                if (envelopeMaxRz != null)
                    MaxAsPos = Math.Round(envelopeMaxRz.Values.ToList().Max(), 2);
                else
                    MaxAsPos = 0;
                if (envelopeMinRz != null)
                    MaxAsNeg = Math.Abs(Math.Round(envelopeMinRz.Values.ToList().Min(), 2));
                else
                    MaxAsNeg = 0;
            }

            Rectangle bounds = printDocument1.DefaultPageSettings.Bounds;
            e.Graphics.DrawString("SAP2000 Shell Designer - Ver. " + Application.ProductVersion,
                                  new Font(FontFamily.GenericSansSerif, 9f, FontStyle.Bold), Brushes.Black,
                                  -50, -50);

            e.Graphics.DrawString("Structure: ",
                                  new Font(FontFamily.GenericSansSerif, 9f, FontStyle.Bold), Brushes.DarkRed,
                                  0, 0);
            e.Graphics.DrawString(txtStructureName.Text,
                                  new Font(FontFamily.GenericSansSerif, 11f, FontStyle.Bold), Brushes.DarkGreen,
                                  100, 0);
            e.Graphics.DrawString("Shell: ",
                                  new Font(FontFamily.GenericSansSerif, 9f, FontStyle.Bold), Brushes.DarkRed,
                                  0, 20);
            e.Graphics.DrawString(txtWallName.Text,
                                  new Font(FontFamily.GenericSansSerif, 11f, FontStyle.Bold), Brushes.DarkGreen,
                                  100, 20);
            e.Graphics.DrawString("Strip: ",
                                  new Font(FontFamily.GenericSansSerif, 9f, FontStyle.Bold), Brushes.DarkRed,
                                  0, 40);
            e.Graphics.DrawString(txtStripName.Text,
                                  new Font(FontFamily.GenericSansSerif, 11f, FontStyle.Bold), Brushes.DarkGreen,
                                  100, 40);
            e.Graphics.DrawString("Width: ",
                                  new Font(FontFamily.GenericSansSerif, 9f, FontStyle.Bold), Brushes.DarkRed,
                                  0, 60);
            e.Graphics.DrawString(Math.Round(be, 2).ToString(),
                                  new Font(FontFamily.GenericSansSerif, 11f, FontStyle.Bold), Brushes.DarkGreen,
                                  100, 60);
            e.Graphics.DrawString("Thickness: ",
                                  new Font(FontFamily.GenericSansSerif, 9f, FontStyle.Bold), Brushes.DarkRed,
                                  0, 80);
            e.Graphics.DrawString(Math.Round(th, 2).ToString(),
                                  new Font(FontFamily.GenericSansSerif, 11f, FontStyle.Bold), Brushes.DarkGreen,
                                  100, 80);
            chart1.Printing.PrintPaint(e.Graphics,
                                       new Rectangle(new System.Drawing.Point(0, 100),
                                                     new Size((int)(bounds.Width * 0.80), (int)(bounds.Width * 0.4))));

            e.Graphics.DrawString(GetACIDescription(),
                      new Font(FontFamily.GenericSansSerif, 9f, FontStyle.Bold), Brushes.Black,
                      0, (int)(bounds.Width * 0.4 + 80));

            e.Graphics.DrawString("Max. As+ (Design Loads) : ",
                                  new Font(FontFamily.GenericSansSerif, 9f, FontStyle.Bold), Brushes.DarkRed,
                                  0, (int)(bounds.Width * 0.4 + 190));
            e.Graphics.DrawString(MaxAsPos + " cm2",
                                  new Font(FontFamily.GenericSansSerif, 11f, FontStyle.Bold), Brushes.Black,
                                  220, (int)(bounds.Width * 0.4 + 190));
            e.Graphics.DrawString("Max. As- (Design Loads) : ",
                                  new Font(FontFamily.GenericSansSerif, 9f, FontStyle.Bold), Brushes.DarkRed,
                                  330, (int)(bounds.Width * 0.4 + 190));
            e.Graphics.DrawString(MaxAsNeg + " cm2",
                                  new Font(FontFamily.GenericSansSerif, 11f, FontStyle.Bold), Brushes.Black,
                                  550, (int)(bounds.Width * 0.4 + 190));

            e.Graphics.DrawString("Max. M+ (Service Loads) : ",
                                  new Font(FontFamily.GenericSansSerif, 9f, FontStyle.Bold), Brushes.DarkRed,
                                  0, (int)(bounds.Width * 0.4 + 210));
            e.Graphics.DrawString(MaxMomentServicePos + " t.m",
                                  new Font(FontFamily.GenericSansSerif, 11f, FontStyle.Bold), Brushes.Black,
                                  220, (int)(bounds.Width * 0.4 + 210));
            e.Graphics.DrawString("Max. M- (Service Loads) : ",
                                  new Font(FontFamily.GenericSansSerif, 9f, FontStyle.Bold), Brushes.DarkRed,
                                  330, (int)(bounds.Width * 0.4 + 210));
            e.Graphics.DrawString(MaxMomentServiceNeg + " t.m",
                                  new Font(FontFamily.GenericSansSerif, 11f, FontStyle.Bold), Brushes.Black,
                                  550, (int)(bounds.Width * 0.4 + 210));

            DrawTable(
                MakeDataTable(MaxAsPos, MaxMomentServicePos, Convert.ToDouble(txtCoverPos.Text.Trim()),
                              Convert.ToInt16(txtPosRows.Text.Trim()), Convert.ToDouble(txtMaxSpace.Text.Trim())),
                e.Graphics, 26, 62, new System.Drawing.Point(0, (int)(bounds.Width * 0.4 + 240)));
            DrawTable(
                MakeDataTable(MaxAsNeg, MaxMomentServiceNeg, Convert.ToDouble(txtCoverNeg.Text.Trim()),
                              Convert.ToInt16(txtNegRows.Text.Trim()), Convert.ToDouble(txtMaxSpace.Text.Trim())),
                e.Graphics, 26, 62, new System.Drawing.Point(330, (int)(bounds.Width * 0.4 + 240)));
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            printDocument1.OriginAtMargins = true;
            printPreviewDialog1.ShowDialog();
            //            printDocument1.Print();
        }
        private string GetACIDescription()
        {
            string limitCons = "";
            if (LimitationForTempRebar)
            {
                limitCons = "(Limitations Considered)";
            }
            string line1 = "Design Code/Method: ACI / Ultimate Strength";
            string line2 = "Fy= " + Convert.ToDouble(txtFy.Text) / 10 + " Kg / cm2";
            string line3 = "Fc= " + Convert.ToDouble(txtFc.Text) / 10 + " Kg / cm2";
            string line4 = "Shrinkage/Temprature Rebar: " + cmbTempReinbar.Text + limitCons;
            string line5 = "Cover (Pos): " + Convert.ToDouble(txtCoverPos.Text) * 100 + " cm";
            string line6 = "Cover (Neg): " + Convert.ToDouble(txtCoverNeg.Text) * 100 + " cm";
            string line7 = "Max. Space Between Bars: " + Convert.ToDouble(txtMaxSpace.Text) * 100 + " cm";
            return line1 + "\r\n" + line2 + "\r\n" + line3 + "\r\n" + line4 + "\r\n" + line5 + "\r\n" + line6 + "\r\n" +
                   line7;
        }
        private double AciRectBeamDesign(double m, double b, double h, double c, double fc, double fy, int tempRebarInd, bool vertical)
        {
            fc = .1 * fc;
            fy = .1 * fy;
            h = h * 100;
            c = c * 100;
            b = b * 100;
            int sgn = Math.Sign(m);
            double mn = m / 0.9;
            double asr = Math.Abs(((((0.85 * 0.6) * (fc / 10)) * (b * 10) * ((h-c) * 10)) / (0.85 * fy / 10)) * (1 - Math.Sqrt(1.0 - (2.0 * (Math.Abs(m) * 10000000) / (0.85 * (0.6 * fc / 10) * (b * 10) * Math.Pow(((h-c) * 10.0), 2)))))) / 100;
            //double asr =
            //    Math.Abs(b * (h - c) / (fy / (0.85 * fc)) * (1 - Math.Sqrt(1 - 2.35 * mn * 100000 / (fc * b * Math.Pow(h - c, 2)))));
            double minCalcAs = 0;
            minCalcAs = 14 / fy * b * (h - c);
            double as133 = 4.0 / 3 * asr;
            double finalAsCalc = 0;
            double tempAs = 0;
            if (asr > minCalcAs)
            {
                finalAsCalc = asr;
            }
            else if (as133 > minCalcAs)
            {
                finalAsCalc = minCalcAs;

            }
            else
            {
                finalAsCalc = as133;
            }
            LimitationForTempRebar = false;
            tempAs = GetTempratureArea(b, h, vertical);
            if (Math.Round(finalAsCalc, 2) == 0)
            {
                int s = sgn;
            }
            if (finalAsCalc < tempAs)
            {
                return tempAs * sgn;
            }
            return finalAsCalc * sgn;
        }
        private void DrawTable(DataTable dataTable, System.Drawing.Graphics g, int rowHeight, int colWidth, System.Drawing.Point startPoint)
        {
            var pen = new Pen(Color.Black);
            for (int i = 0; i < dataTable.Columns.Count + 1; i++)
            {
                var ps = new System.Drawing.Point((int)(startPoint.X + i * colWidth), startPoint.Y);
                var pe = new System.Drawing.Point((int)(startPoint.X + i * colWidth), startPoint.Y + rowHeight * (dataTable.Rows.Count));
                g.DrawLine(pen, ps, pe);
            }
            for (int i = 0; i < dataTable.Rows.Count + 1; i++)
            {
                var ps = new System.Drawing.Point((int)(startPoint.X), startPoint.Y + i * rowHeight);
                var pe = new System.Drawing.Point((int)(startPoint.X + colWidth * dataTable.Columns.Count), startPoint.Y + i * rowHeight);
                g.DrawLine(pen, ps, pe);
            }
            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                string content = dataTable.Rows[0][i].ToString();
                var p = new System.Drawing.Point((int)(startPoint.X + i * colWidth), startPoint.Y + (0) * rowHeight);
                var font = new Font(FontFamily.GenericSansSerif, 9, FontStyle.Regular);
                g.DrawString(content, font, Brushes.Red, p);
            }
            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                for (int j = 1; j < dataTable.Rows.Count; j++)
                {
                    string content = dataTable.Rows[j][i].ToString();
                    var p = new System.Drawing.Point((int)(startPoint.X + i * colWidth), startPoint.Y + (j) * rowHeight);
                    var font = new Font(FontFamily.GenericSansSerif, 9, FontStyle.Regular);
                    g.DrawString(content, font, Brushes.Black, p);
                }
            }
        }
        private bool LimitationForTempRebar;
        private DataTable MakeDataTable(double asMax, double mMax, double cover, int rows, double maxSpace)
        {
            var myDataTable = new DataTable();
            myDataTable.Columns.Add("colRebar");
            myDataTable.Columns.Add("colSpace");
            myDataTable.Columns.Add("colDc");
            myDataTable.Columns.Add("colFs");
            myDataTable.Columns.Add("colCrack");
            myDataTable.Rows.Add();
            myDataTable.Rows[0]["colRebar"] = "Bar(mm)";
            myDataTable.Rows[0]["colSpace"] = " s (cm)";
            myDataTable.Rows[0]["colDc"] = " dc (cm)";
            myDataTable.Rows[0]["colFs"] = "fs(kg/cm2)";
            myDataTable.Rows[0]["colCrack"] = " w (mm)";
            double[] bars = new double[] { 12, 14, 16, 18, 20, 22, 24, 25, 26, 28, 30, 32, 34, 36, 38 };
            for (int i = 0; i < bars.Length; i++)
            {
                double barDim = (double)bars.GetValue(i);
                double barArea = Math.Pow(barDim, 2) * Math.PI / 4;
                int barNumber = Convert.ToInt32(Math.Truncate((asMax / rows * 100 / barArea))) + 1;
                double space = CurrentStripWidth * 1000;
                if (barNumber > 2)
                {
                    space = CurrentStripWidth / (barNumber - 1) * 1000;
                }
                if (space > maxSpace * 1000) space = maxSpace * 1000;
                double dc = cover * 1000 + barDim / 2;
                double spaceArea = space * (dc) * 2;
                double fs = mMax * 10000000 / (asMax * 100 * (th * 1000 - dc));

                double w = 13 * fs * Math.Pow((double)(dc * spaceArea), (double)(0.3333));
                w /= 1000000;

                myDataTable.Rows.Add();
                myDataTable.Rows[i + 1]["colRebar"] = bars.GetValue(i).ToString();
                if (rows > 1)
                {
                    myDataTable.Rows[i + 1]["colSpace"] = Math.Round(space / 10, 1) + " (" + rows + ")";
                }
                else
                {
                    myDataTable.Rows[i + 1]["colSpace"] = Math.Round(space / 10, 1);
                }
                myDataTable.Rows[i + 1]["colDc"] = Math.Round(dc / 10, 1);
                myDataTable.Rows[i + 1]["colFs"] = Math.Round(fs * 10, 1);
                myDataTable.Rows[i + 1]["colCrack"] = Math.Round(w, 3);
            }
            return myDataTable;
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            ReDesign();
            UpdateChart();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var str = new SelectionData
            {
                StripName = txtStripName.Text.Trim(),
                SelElementNames = selList,
                DesignDefinition = new DesignDefinition()
                {
                    DesignMethod = cmbDesignMethod.SelectedIndex,
                    Fc = Convert.ToDouble(txtFc.Text.Trim()),
                    Fs = Convert.ToDouble(txtFs.Text.Trim()),
                    Fy = Convert.ToDouble(txtFy.Text.Trim()),
                    MaxReinSpace = Convert.ToDouble(txtMaxSpace.Text.Trim()),
                    NegCover = Convert.ToDouble(txtCoverNeg.Text.Trim()),
                    PosCover = Convert.ToDouble(txtCoverPos.Text.Trim()),
                    NegRowNum = Convert.ToInt16(txtNegRows.Text.Trim()),
                    PosRowNum = Convert.ToInt16(txtPosRows.Text.Trim()),
                    TempRein = cmbTempReinbar.SelectedIndex
                },
                OutputDefinition = new OutputDefinition()
                {
                    Combo = cmbComboNames.SelectedIndex,
                    CrackCombos = GetIndecies(chkComboListCrack),
                    StrengthCombos = GetIndecies(chkCombosListDesign),
                    StripMomentType = cmbStripMomentType.SelectedIndex,
                    StripOutputType = cmbStripOutputType.SelectedIndex
                }
            };

            List<SelectionData> dic = new List<SelectionData>();
            dic.Add(str);
            string s = GenericSerializer.Serialize(dic);

        }
        private bool ListContainsStripName(List<SelectionData> listData, string shellStripName, ref SelectionData listItem)
        {
            foreach (SelectionData selectionData in listData)
            {
                if (selectionData.ShellName + "-" + selectionData.StripName == shellStripName)
                {
                    listItem = selectionData;
                    return true;
                }
            }
            return false;
        }
        private string fPath;

        private void tsmOpenState_Click(object sender, EventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "DAT Files|*.dat";
            if (openFileDialog.ShowDialog() != System.Windows.Forms.DialogResult.Abort)
            {
                fPath = openFileDialog.FileName;
                SelStrips = (List<SelectionData>)HistoryMaker.LoadData(fPath);
                SelectionDone = true;
                RefreshListBox();
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }


        private List<SelectionData> SelStrips;
        private int[] GetIndecies(CheckedListBox lstBox)
        {
            int[] allIndices = new int[lstBox.CheckedIndices.Count];
            int n = 0;
            foreach (int index in lstBox.CheckedIndices)
            {
                allIndices.SetValue(index, n);
                n++;
            }
            return allIndices;
        }
        private void button3_Click(object sender, EventArgs e)
        {
            if (SelStrips == null) SelStrips = new List<SelectionData>();
            var str = new SelectionData
            {
                StuctureName = txtStructureName.Text.Trim(),
                ShellName = txtWallName.Text.Trim(),
                StripName = txtStripName.Text.Trim(),
                SelElementNames = selList,
                DesignDefinition = new DesignDefinition()
                {
                    DesignMethod = cmbDesignMethod.SelectedIndex,
                    Fc = Convert.ToDouble(txtFc.Text.Trim()),
                    Fs = Convert.ToDouble(txtFs.Text.Trim()),
                    Fy = Convert.ToDouble(txtFy.Text.Trim()),
                    MaxReinSpace = Convert.ToDouble(txtMaxSpace.Text.Trim()),
                    NegCover = Convert.ToDouble(txtCoverNeg.Text.Trim()),
                    PosCover = Convert.ToDouble(txtCoverPos.Text.Trim()),
                    NegRowNum = Convert.ToInt16(txtNegRows.Text.Trim()),
                    PosRowNum = Convert.ToInt16(txtPosRows.Text.Trim()),
                    TempRein = cmbTempReinbar.SelectedIndex
                },
                OutputDefinition = new OutputDefinition()
                {
                    Combo = cmbComboNames.SelectedIndex,
                    CrackCombos = GetIndecies(chkComboListCrack),
                    StrengthCombos = GetIndecies(chkCombosListDesign),
                    StripMomentType = cmbStripMomentType.SelectedIndex,
                    StripOutputType = cmbStripOutputType.SelectedIndex
                }
            };

            SelectionData selData = new SelectionData();
            if (ListContainsStripName(SelStrips, str.ShellName + "-" + str.StripName, ref selData))
            {
                DialogResult result = MessageBox.Show("A Strip with this name exists!\r\nDo you want to replace it?", "Sap2000 Shell Designer", MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {
                    SelStrips.Remove(selData);
                    SelStrips.Add(str);
                }
            }
            else
            {
                SelStrips.Add(str);
            }
            RefreshListBox();
        }

        private void RefreshListBox()
        {
            listBox1.Items.Clear();
            foreach (SelectionData selStrip in SelStrips)
            {
                listBox1.Items.Add(selStrip.ShellName + "-" + selStrip.StripName);
            }
        }

        private void tsmAddToCurrent_Click(object sender, EventArgs e)
        {

        }

        private void tsmSaveAs_Click(object sender, EventArgs e)
        {
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "DAT Files|*.dat";
            if (saveFileDialog.ShowDialog() != System.Windows.Forms.DialogResult.Cancel)
            {
                fPath = saveFileDialog.FileName;
                HistoryMaker.SaveSelectionToFile(fPath, SelStrips);
            }
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            lastSelectedIndex = listBox1.SelectedIndex;
            SwitchStrip();
        }
        private void SwitchStrip()
        {
            var selData = SelStrips[lastSelectedIndex];

            selList = selData.SelElementNames;
            SelectionDone = true;

            txtStructureName.Text = selData.StuctureName;
            txtWallName.Text = selData.ShellName;
            txtStripName.Text = selData.StripName;

            txtCoverPos.Text = selData.DesignDefinition.PosCover.ToString();
            txtCoverNeg.Text = selData.DesignDefinition.NegCover.ToString();
            cmbDesignMethod.SelectedIndex = selData.DesignDefinition.DesignMethod;

            txtFy.Text = selData.DesignDefinition.Fy.ToString();
            txtFc.Text = selData.DesignDefinition.Fc.ToString();
            txtFs.Text = selData.DesignDefinition.Fs.ToString();
            cmbTempReinbar.SelectedIndex = selData.DesignDefinition.TempRein;
            txtPosRows.Text = selData.DesignDefinition.PosRowNum.ToString();
            txtNegRows.Text = selData.DesignDefinition.NegRowNum.ToString();
            txtMaxSpace.Text = selData.DesignDefinition.MaxReinSpace.ToString();
            cmbComboNames.SelectedIndex = selData.OutputDefinition.Combo;
            cmbStripMomentType.SelectedIndex = selData.OutputDefinition.StripMomentType;
            cmbStripOutputType.SelectedIndex = selData.OutputDefinition.StripOutputType;

            UncheckedAll(chkComboListCrack);
            UncheckedAll(chkCombosListDesign);

            for (int i = 0; i < selData.OutputDefinition.CrackCombos.Length; i++)
            {
                chkComboListCrack.SetItemChecked((int)selData.OutputDefinition.CrackCombos.GetValue(i), true);
            }
            for (int i = 0; i < selData.OutputDefinition.StrengthCombos.Length; i++)
            {
                chkCombosListDesign.SetItemChecked((int)selData.OutputDefinition.StrengthCombos.GetValue(i), true);
            }
        }

        private void UncheckedAll(CheckedListBox list)
        {
            for (int i = 0; i < list.Items.Count; i++)
            {
                list.SetItemChecked(i, false);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            UpdateSeleStrip();
        }

        private int lastSelectedIndex;
        private void UpdateSeleStrip()
        {
            var selData = new SelectionData();
            selData.SelElementNames = selList;

            selData.StuctureName = txtStructureName.Text;
            selData.ShellName = txtWallName.Text;
            selData.StripName = txtStripName.Text;

            selData.DesignDefinition.PosCover = Convert.ToDouble(txtCoverPos.Text);
            selData.DesignDefinition.NegCover = Convert.ToDouble(txtCoverNeg.Text);
            selData.DesignDefinition.DesignMethod = cmbDesignMethod.SelectedIndex;

            selData.DesignDefinition.Fy = Convert.ToDouble(txtFy.Text);
            selData.DesignDefinition.Fc = Convert.ToDouble(txtFc.Text);
            selData.DesignDefinition.Fs = Convert.ToDouble(txtFs.Text);

            selData.DesignDefinition.TempRein = cmbTempReinbar.SelectedIndex;
            selData.DesignDefinition.PosRowNum = Convert.ToInt32(txtPosRows.Text);
            selData.DesignDefinition.NegRowNum = Convert.ToInt32(txtNegRows.Text);
            selData.DesignDefinition.MaxReinSpace = Convert.ToDouble(txtMaxSpace.Text);
            selData.OutputDefinition.Combo = cmbComboNames.SelectedIndex;
            selData.OutputDefinition.StripMomentType = cmbStripMomentType.SelectedIndex;
            selData.OutputDefinition.StripOutputType = cmbStripOutputType.SelectedIndex;

            selData.OutputDefinition.CrackCombos = GetIndecies(chkComboListCrack);
            selData.OutputDefinition.StrengthCombos = GetIndecies(chkCombosListDesign);
            SelStrips[lastSelectedIndex] = selData;
            RefreshListBox();
            listBox1.SelectedIndex = lastSelectedIndex;
        }
        private void btnDelStrip_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Do You Really Want To Delete This Strip?", "SAP2000 Shell Designer"
                , MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2, MessageBoxOptions.DefaultDesktopOnly) == DialogResult.Yes)
                if (listBox1.Items.Count > 0)
                {
                    SelStrips.RemoveAt(listBox1.SelectedIndex);
                    RefreshListBox();
                }
            if (listBox1.Items.Count > 0) listBox1.SelectedIndex = 0;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            FillObjects();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < listBox1.Items.Count; i++)
            {
                listBox1.SelectedIndex = i;
                FillObjects();
                ReDesign();
                UpdateChartForEnvelope();
                printDocument1.OriginAtMargins = true;
                printDocument1.Print();
            }

        }

        private void button7_Click(object sender, EventArgs e)
        {
            sapShell.SelectElements(selList);
        }
    }
}
