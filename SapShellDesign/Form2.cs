using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SapShellDesign
{
    public partial class Form2 : Form
    {
        private readonly SapShell _sapshell;
        public Form2(Dictionary<string, List<ShellElement>> xyplane, Dictionary<string, List<ShellElement>> xzplane,Dictionary<string, List<ShellElement>> yzplane, SapShell sapShell)
        {
            XYPlaneElements = xyplane;
            XZPlaneElements = xzplane;
            YZPlaneElements = yzplane;
            _sapshell = sapShell;
            InitializeComponent();
        }
        private Dictionary<string, List<ShellElement>> XYPlaneElements;
        private Dictionary<string, List<ShellElement>> XZPlaneElements;
        private Dictionary<string, List<ShellElement>> YZPlaneElements;
        private void Form2_Load(object sender, EventArgs e)
        {
            treeView1.Nodes.Clear();
            treeView1.Nodes.Add("xyPlanes", "XY-Planes");
            treeView1.Nodes.Add("xzPlanes", "XZ-Planes");
            treeView1.Nodes.Add("yzPlanes", "YZ-Planes");
            int i = 0;
            foreach (KeyValuePair<string, List<ShellElement>> planeElement in XYPlaneElements)
            {
                i++;
                treeView1.Nodes["xyPlanes"].Nodes.Add("xyPlaneName" + i, "Plane-" + i);
                treeView1.Nodes["xyPlanes"].Nodes["xyPlaneName" + i].Nodes.Add("xyPlaneKey", "Key");                
                treeView1.Nodes["xyPlanes"].Nodes["xyPlaneName" + i].Nodes["xyPlaneKey"].Nodes.Add(planeElement.Key, planeElement.Key);
            }
            i = 0;
            foreach (KeyValuePair<string, List<ShellElement>> planeElement in XZPlaneElements)
            {
                i++;
                treeView1.Nodes["xzPlanes"].Nodes.Add("xzPlaneName" + i, "Plane-" + i);
                treeView1.Nodes["xzPlanes"].Nodes["xzPlaneName" + i].Nodes.Add("xzPlaneKey", "Key");
                treeView1.Nodes["xzPlanes"].Nodes["xzPlaneName" + i].Nodes["xzPlaneKey"].Nodes.Add(planeElement.Key, planeElement.Key);
            }
            i = 0;
            foreach (KeyValuePair<string, List<ShellElement>> planeElement in YZPlaneElements)
            {
                i++;
                treeView1.Nodes["yzPlanes"].Nodes.Add("yzPlaneName" + i, "Plane-" + i);
                treeView1.Nodes["yzPlanes"].Nodes["yzPlaneName" + i].Nodes.Add("yzPlaneKey", "Key");
                treeView1.Nodes["yzPlanes"].Nodes["yzPlaneName" + i].Nodes["yzPlaneKey"].Nodes.Add(planeElement.Key, planeElement.Key);
            }
        }
        private List<ShellElement> FindElementsInSamePlane(string planeKey)
        {
            double A = Convert.ToDouble(planeKey.Split(',')[0].Trim().Split('=')[1].Trim());
            double B = Convert.ToDouble(planeKey.Split(',')[1].Trim().Split('=')[1].Trim());
            double C = Convert.ToDouble(planeKey.Split(',')[2].Trim().Split('=')[1].Trim());
            double D = Convert.ToDouble(planeKey.Split(',')[3].Trim().Split('=')[1].Trim());
            var myList = new List<ShellElement>();
            foreach (KeyValuePair<string, List<ShellElement>> element in XYPlaneElements)
            {
                if (VectorOperations.PlaneMatches(A, B, C, D, element.Value[0].GetPlane()))
                {
                    foreach (ShellElement shellElement in element.Value)
                    {
                        myList.Add(shellElement);
                    }
                    return myList;
                }
            }           
            foreach (KeyValuePair<string, List<ShellElement>> element in XZPlaneElements)
            {
                if (VectorOperations.PlaneMatches(A, B, C, D, element.Value[0].GetPlane()))
                {
                    foreach (ShellElement shellElement in element.Value)
                    {
                        myList.Add(shellElement);
                    }
                    return myList;
                }
            }
            foreach (KeyValuePair<string, List<ShellElement>> element in YZPlaneElements)
            {
                if (VectorOperations.PlaneMatches(A, B, C, D, element.Value[0].GetPlane()))
                {
                    foreach (ShellElement shellElement in element.Value)
                    {
                        myList.Add(shellElement);
                    }
                    return myList;
                }
            }

            return myList;
        }

        private void treeView1_DoubleClick(object sender, EventArgs e)
        {
            DrawPlane(pictureBox1.CreateGraphics());
        }
        private void pictureBox1_Paint(object sender, PaintEventArgs e)
        {
            DrawPlane(pictureBox1.CreateGraphics());
        }

        private List<ShellElement> CurrentElements;
        private List<ShellElement> CurrentStrips;
        private void DrawPlane(System.Drawing.Graphics g)
        {
            g.Clear(Color.Black);

            int magnF = 10;
            try
            {
                //pictureBox1.Scale(10, 10);
                CurrentElements = FindElementsInSamePlane(treeView1.SelectedNode.Text);
                foreach (ShellElement element in CurrentElements)
                {
                    System.Drawing.Point p1 = new System.Drawing.Point((int)(element.Points.Values.ToList()[0].X * magnF), (int)(element.Points.Values.ToList()[0].Y * magnF));
                    System.Drawing.Point p2 = new System.Drawing.Point((int)(element.Points.Values.ToList()[1].X * magnF), (int)(element.Points.Values.ToList()[1].Y * magnF));
                    System.Drawing.Point p3 = new System.Drawing.Point((int)(element.Points.Values.ToList()[2].X * magnF), (int)(element.Points.Values.ToList()[2].Y * magnF));
                    System.Drawing.Point p4 = new System.Drawing.Point((int)(element.Points.Values.ToList()[3].X * magnF), (int)(element.Points.Values.ToList()[3].Y * magnF));
                    //g.DrawLine(new Pen(Color.Black), p1, p2);
                    //g.DrawLine(new Pen(Color.Black), p2, p3);
                    //g.DrawLine(new Pen(Color.Black), p3, p4);
                    //g.DrawLine(new Pen(Color.Black), p4, p1);
                    g.DrawRectangle(new Pen(Color.Red), new Rectangle(p1, new Size(p3)));                    
                }
            }
            catch
            {
            }
        }
        private void DrawStrip(System.Drawing.Graphics g)
        {
            g.Clear(Color.Black);
            int magnF = 10;
            try
            {
                //pictureBox1.Scale(10, 10);
                foreach (ShellElement element in CurrentStrips)
                {
                    System.Drawing.Point p1 = new System.Drawing.Point((int)(element.Points.Values.ToList()[0].X * magnF), (int)(element.Points.Values.ToList()[0].Y * magnF));
                    System.Drawing.Point p2 = new System.Drawing.Point((int)(element.Points.Values.ToList()[1].X * magnF), (int)(element.Points.Values.ToList()[1].Y * magnF));
                    System.Drawing.Point p3 = new System.Drawing.Point((int)(element.Points.Values.ToList()[2].X * magnF), (int)(element.Points.Values.ToList()[2].Y * magnF));
                    System.Drawing.Point p4 = new System.Drawing.Point((int)(element.Points.Values.ToList()[3].X * magnF), (int)(element.Points.Values.ToList()[3].Y * magnF));
                    //g.DrawLine(new Pen(Color.Black), p1, p2);
                    //g.DrawLine(new Pen(Color.Black), p2, p3);
                    //g.DrawLine(new Pen(Color.Black), p3, p4);
                    //g.DrawLine(new Pen(Color.Black), p4, p1);
                    g.FillRectangle(Brushes.Gray, new Rectangle(p1, new Size(p3)));
                }
            }
            catch
            {
            }
        }
        
        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            DrawPlane(e.Graphics);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            printDocument1.OriginAtMargins = true;
            printPreviewDialog1.ShowDialog();
            //printDocument1.Print();
        }

        private void pictureBox1_Validated(object sender, EventArgs e)
        {
            DrawPlane(pictureBox1.CreateGraphics());
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            List<string> selShells=_sapshell.GetSelectedElements();
            CurrentStrips=new List<ShellElement>();
            foreach (string selShellName in selShells)
            {
                for (int i = 0; i < CurrentElements.Count; i++)
                {
                    if (CurrentElements[i].Name==selShellName)
                    {
                        CurrentStrips.Add(CurrentElements[i]);
                    }
                }
            }
            if (CurrentStrips!=null)
            {
                DrawStrip(pictureBox1.CreateGraphics());
            }
        }    
    }
}
