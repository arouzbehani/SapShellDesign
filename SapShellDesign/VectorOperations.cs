using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SapShellDesign
{
    [Serializable]
    public struct XYZPoint
    {
        public double X;
        public double Y;
        public double Z;
    }
    [Serializable]
    public class Plane
    {
        public Plane(XYZPoint p1,XYZPoint p2, XYZPoint p3)
        {
            double x1 = p1.X;
            double y1 = p1.Y;
            double z1 = p1.Z;

            double x2 = p2.X;
            double y2 = p2.Y;
            double z2 = p2.Z;

            double x3 = p3.X;
            double y3 = p3.Y;
            double z3 = p3.Z;

            A = y1 * (z2 - z3) + y2 * (z3 - z1) + y3 * (z1 - z2);
            B = z1*(x2 - x3) + z2*(x3 - x1) + z3*(x1 - x2);
            C = x1*(y2 - y3) + x2*(y3 - y1) + x3*(y1 - y2);
            D =-1* (x1*(y2*z3 - y3*z2) + x2*(y3*z1 - y1*z3) + x3*(y1*z2 - y2*z1));
        }
        public double A;
        public double B;
        public double C;
        public double D;
        public string GetPlaneScript()
        {
            return string.Format("A={0}, B={1}, C={2}, D={3}", Math.Round(A, 1), Math.Round(B, 1),
                                 Math.Round(C, 1), Math.Round(D, 1));
        }
    }

    [Serializable]
    public struct Vector
    {
        public XYZPoint Pi;
        public XYZPoint Pj;
        public double a()
        {
            return Pj.X - Pi.X;
        }
        public double b()
        {
            return Pj.Y - Pi.Y;
        }
        public double c()
        {
            return Pj.Z - Pi.Z;
        }
        public double Length()
        {
            return Math.Sqrt(a()*a() + b()*b() + c()*c());
        }
        public double CosAlpha()
        {
            if (Length() != 0)
                return a() / Length();
            return 0;
        }
        public double CosBeta()
        {
            if (Length() != 0)
                return b() / Length();
            return 0;
        }
        public double CosGama()
        {
            if (Length() != 0)
                return c()/Length();
            return 0;
        }
        public Vector Unique()
        {
            var uv = new Vector {Pi = Pi, Pj = new XYZPoint()};
            if(0!=Length())
            {                
                uv.Pj.X = Pi.X + a()/Length();
                uv.Pj.Y = Pi.Y + b()/Length();
                uv.Pj.Z = Pi.Z + c()/Length();
            }
            else
            {
                uv.Pj = Pi;
            }
            return uv;
        }
    }

    [Serializable]
    class VectorOperations
    {
        public static Vector GetNormalVector (Vector v1 , Vector v2)
        {
            var nv = new Vector {Pi = v1.Pi};
            double a = v1.b()*v2.c() - v1.c()*v2.b();
            double b = v1.c()*v2.a() - v1.a()*v2.c();
            double c = v1.a()*v2.b() - v1.b()*v2.a();
            nv.Pj = new XYZPoint() {X = nv.Pi.X + a, Y = nv.Pi.Y + b, Z = nv.Pi.Z + c};
            return nv;
        }
        public static bool VectorsAreParallel(Vector v1, Vector v2)
        {
            if (GetNormalVector(v1, v2).Length() == 0)
                return true;
            return false;
        }
        public static bool IsInPlane(Plane plane,XYZPoint point)
        {
            if (0!=Math.Round(plane.A*point.X+plane.B*point.Y+plane.C*point.Z+plane.D,2))
            {
                return false;
            }
            return true;
        }
        public static bool PlaneMatches(double A,double B,double C,double D, Plane plane)
        {
            if (Math.Round(A,1)==Math.Round(plane.A,1))
            {
                if (Math.Round(B, 1) == Math.Round(plane.B, 1))
                {
                    if (Math.Round(C, 1) == Math.Round(plane.C, 1))
                    {
                        if (Math.Round(D, 1) == Math.Round(plane.D, 1))
                        {
                            return true;
                        }
                    }
                }                
            }
            return false;
        }
    }
}
