using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System.IO;
using OfficeOpenXml;

[assembly: CommandClass(typeof(MyFirstProject1.Class1))]

namespace MyFirstProject1
{

    public class Class1
    {
        static Dictionary<int, int> bc = new Dictionary<int, int>();//表格里的长度统计
        static Dictionary<int, int> gc = new Dictionary<int, int>();//余料统计
        static ArrayList globalChoose;
        public static void DrawBorder(Hatch hat, Transaction trans, BlockTableRecord btr, int numSample)
        {
            //取得边界数  

        }


        [CommandMethod("AddLine")]
        public static void AddLineCmd()
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Transaction acTrans = acCurDb.TransactionManager.StartTransaction();
            Database db = acDoc.Database;//获得当前工作空间的数据库
            BlockTable bt = (BlockTable)db.BlockTableId.Open(OpenMode.ForRead); //获得块表
            BlockTableRecord btr = (BlockTableRecord)bt[BlockTableRecord.ModelSpace].Open(OpenMode.ForRead);    //获得模型空间的块表记录
            int cnt = 0;
            Console.WriteLine("\n对象：");
            foreach (ObjectId acObjId in btr)
            {
                Entity acEnt = acTrans.GetObject(acObjId, OpenMode.ForWrite) as Entity;
            }
            Line line = new Line(new Autodesk.AutoCAD.Geometry.Point3d(0, 0, 0), new Autodesk.AutoCAD.Geometry.Point3d(914837.7813, 534890.7717, 0));
            try
            {
                btr.AppendEntity(line);  //将直线添加到模型空间中
                line.Close();//关闭直线
            }
            finally
            {
                btr.Close();//关闭块表记录
                bt.Close();//关闭块表
            }
        }

        public static void cutLine(Point2d a, Point2d b, ref Database acCurDb, double textposition, Transaction acTrans, ref BlockTableRecord acBlkTblRec)
        {

            if (a.GetDistanceTo(b) > 80000)
            {
                return;
            }
            Point2d ori = new Point2d();




            while (a.GetDistanceTo(b) - 600 > 200 || Math.Abs(a.GetDistanceTo(b) - 600) < 2)
            {

                if (Math.Abs(a.X - b.X) < 1)
                {//竖
                    if (a.Y > b.Y)//上向下
                    {
                        ori = a;
                        a = new Point2d(a.X, a.Y - 600);
                    }
                    else//下向上
                    {
                        ori = a;
                        a = new Point2d(a.X, a.Y + 600);
                    }

                }
                else//横
                {
                    if (a.X < b.X)//左向右
                    {
                        ori = a;
                        a = new Point2d(a.X + 600, a.Y);
                    }
                    else//右向左
                    {
                        ori = a;
                        a = new Point2d(a.X - 600, a.Y);
                    }
                }
                AlignedDimension acRotDim = new AlignedDimension();
                acRotDim.DimensionStyle = acCurDb.Dimstyle;
                acRotDim.Dimclrd = Autodesk.AutoCAD.Colors.Color.FromRgb(124, 252, 0);
                acRotDim.Dimclrt = Autodesk.AutoCAD.Colors.Color.FromRgb(124, 252, 0);
                acRotDim.Dimclre = Autodesk.AutoCAD.Colors.Color.FromRgb(124, 252, 0);


                acRotDim.XLine1Point = new Point3d(ori.X, ori.Y, 0);
                acRotDim.XLine2Point = new Point3d(a.X, a.Y, 0);


                if (Math.Abs(ori.X - a.X) < 1.0)//竖
                {
                    acRotDim.DimLinePoint = new Point3d(ori.X + textposition, (a.Y + ori.Y) / 2, 0);
                }
                else
                {
                    acRotDim.DimLinePoint = new Point3d((ori.X + a.X) / 2, a.Y + textposition, 0);
                }
                // 将新对象添加到块表记录 ModelSpace 及事务
                acBlkTblRec.AppendEntity(acRotDim);
                acTrans.AddNewlyCreatedDBObject(acRotDim, true);

                //bc[1200] += 1;
                if (!bc.ContainsKey(600))
                {
                    bc.Add(600, 1);
                }
                else
                {
                    bc[600] += 1;
                }
                if (gc.ContainsKey(600))
                {
                    if(gc[600] > 0)
                    {
                        gc[600]--;
                    }
                    else
                    {
                        gc[600]++;
                    }
                }
                else
                {
                    gc.Add(600, 1);
                }
            }



            //剩1400以下的
            //Dictionary<int, int> gccopy = new Dictionary<int, int>(gc);
            ArrayList ary = new ArrayList();  //剩料的 arraylist格式
            ArrayList choose = new ArrayList(); //用料的 arraylist格式
                                                //先把字典里的数据拿到数组里

            string keyNames = "";

            foreach (int key in gc.Keys)

            {

                if (gc[key] == 0)

                {
                    keyNames += key + ",";
                }

            }

            string[] str_keyNames = keyNames.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (string key in str_keyNames)
            {
                gc.Remove(int.Parse(key));
            }



            foreach (KeyValuePair<int, int> kv in gc)

            {
                int times = kv.Value;

                while (times > 0)
                {
                    ary.Add(kv.Key);
                    times--;
                }

            }
            int leftlen = (int)(a.GetDistanceTo(b) + 1) / 10 * 10;
            bool gd = false;
            //1块能解决
            if (ary.Contains(leftlen) && leftlen > 0)
            {
                gd = true;

                choose.Add(leftlen);
                if (gc[leftlen] == 1)
                {
                    gc.Remove(leftlen);
                }
                else
                {
                    gc[leftlen]--;
                }
                ary.Remove(leftlen);

                //加到材料切割数组里
                if (!bc.ContainsKey((int)((int)a.GetDistanceTo(b) + 1) / 10 * 10))
                {
                    bc.Add(((int)((int)a.GetDistanceTo(b) + 1) / 10) * 10, 1);
                }
                else
                {
                    bc[((int)((int)a.GetDistanceTo(b) + 1) / 10) * 10] += 1;
                }
                leftlen = 0;

            }

            //2块能解决
            if (gd == false && ary.Count >= 2 && leftlen > 0)
            {
                for (int Index1 = 0; Index1 <= ary.Count - 2 && gd == false; Index1++)
                {
                    for (int Index2 = 1; Index2 <= ary.Count - 1 && gd == false; Index2++)
                    {
                        if ((int)ary[Index1] + (int)ary[Index2] == leftlen)
                        {
                            gd = true;

                            choose.Add(ary[Index1]);
                            choose.Add(ary[Index2]);



                            if (gc[(int)ary[Index1]] == 1)
                            {
                                //gc.Remove((int)ary[Index1]);
                                gc[(int)ary[Index1]]--;
                            }
                            else
                            {
                                gc[(int)ary[Index1]]--;
                            }


                            if (!bc.ContainsKey(((int)ary[Index1] + 1) / 10 * 10))
                            {
                                bc.Add((((int)ary[Index1]) / 10) * 10, 1);
                            }
                            else
                            {
                                bc[(((int)ary[Index1] + 1) / 10) * 10] += 1;
                            }
                            /*
                            if (gc[(int)ary[Index2]] == 1)
                            {
                                gc.Remove((int)ary[Index2]);
                            }
                            else
                            {
                                gc[(int)ary[Index2]]--;
                            }
                            */


                            if (!bc.ContainsKey(((int)ary[Index2] + 1) / 10 * 10))
                            {
                                bc.Add((((int)ary[Index2]) / 10) * 10, 1);
                            }
                            else
                            {
                                bc[(((int)ary[Index2] + 1) / 10) * 10] += 1;
                            }
                            leftlen = 0;
                            string tr = "";
                            tr = ary[Index1] + "," + ary[Index2];

                            string[] trs = tr.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                            foreach (string trk in trs)
                            {
                                ary.Remove(int.Parse(trk));
                            }

                            break;
                        }
                    }

                }
            }

            //3块能解决
            if (gd == false && ary.Count >= 3 && leftlen > 0)
            {
                for (int Index1 = 0; Index1 <= ary.Count - 3 && gd == false; Index1++)
                {
                    for (int Index2 = 1; Index2 <= ary.Count - 2 && gd == false; Index2++)
                    {
                        for (int Index3 = 2; Index3 <= ary.Count - 1 && gd == false; Index3++)
                        {
                            if ((int)ary[Index1] + (int)ary[Index2] + (int)ary[Index3] == leftlen)
                            {
                                gd = true;

                                choose.Add(ary[Index1]);
                                choose.Add(ary[Index2]);
                                choose.Add(ary[Index3]);
                                /*
                                if (gc[(int)ary[Index1]] == 1)
                                {
                                    gc.Remove((int)ary[Index1]);
                                }
                                else
                                {
                                    gc[(int)ary[Index1]]--;
                                }
                                */
                                gc[(int)ary[Index1]]--;
                                //ary.Remove((int)ary[Index1]);

                                if (!bc.ContainsKey(((int)ary[Index1] + 1) / 10 * 10))
                                {
                                    bc.Add((((int)ary[Index1]) / 10) * 10, 1);
                                }
                                else
                                {
                                    bc[(((int)ary[Index1] + 1) / 10) * 10] += 1;
                                }
                                /*
                                if (gc[(int)ary[Index2]] == 1)
                                {
                                    gc.Remove((int)ary[Index2]);
                                }
                                else
                                {
                                    gc[(int)ary[Index2]]--;
                                }
                                */
                                gc[(int)ary[Index2]]--;
                                //ary.Remove((int)ary[Index2]);

                                if (!bc.ContainsKey(((int)ary[Index2] + 1) / 10 * 10))
                                {
                                    bc.Add((((int)ary[Index2]) / 10) * 10, 1);
                                }
                                else
                                {
                                    bc[(((int)ary[Index2] + 1) / 10) * 10] += 1;
                                }
                                /*
                                if (gc[(int)ary[Index3]] == 1)
                                {
                                    gc.Remove((int)ary[Index3]);
                                }
                                else
                                {
                                    gc[(int)ary[Index3]]--;
                                }
                                */
                                gc[(int)ary[Index3]]--;
                                //ary.Remove((int)ary[Index3]);
                                leftlen = 0;

                                string tr = "";
                                tr = ary[Index1] + "," + ary[Index2] + "," + ary[Index3];

                                string[] trs = tr.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                                foreach (string trk in trs)
                                {
                                    ary.Remove(int.Parse(trk));
                                }


                                break;
                            }
                        }
                    }
                }
            }
            //4块能解决
            if (gd == false && ary.Count >= 4 && leftlen > 0)
            {
                for (int Index1 = 0; Index1 <= ary.Count - 4 && gd == false; Index1++)
                {
                    for (int Index2 = 1; Index2 <= ary.Count - 3 && gd == false; Index2++)
                    {
                        for (int Index3 = 2; Index3 <= ary.Count - 2 && gd == false; Index3++)
                        {
                            for (int Index4 = 3; Index4 <= ary.Count - 1 && gd == false; Index4++)
                            {
                                if ((int)ary[Index1] + (int)ary[Index2] + (int)ary[Index3] + (int)ary[Index4] == leftlen)
                                {
                                    gd = true;

                                    choose.Add(ary[Index1]);
                                    choose.Add(ary[Index2]);
                                    choose.Add(ary[Index3]);
                                    choose.Add(ary[Index4]);
                                    /*
                                    if (gc[(int)ary[Index1]] == 1)
                                    {
                                        gc.Remove((int)ary[Index1]);
                                    }
                                    else
                                    {
                                        gc[(int)ary[Index1]]--;
                                    }
                                    */
                                    gc[(int)ary[Index1]]--;
                                    //ary.Remove((int)ary[Index1]);

                                    if (!bc.ContainsKey(((int)ary[Index1] + 1) / 10 * 10))
                                    {
                                        bc.Add((((int)ary[Index1]) / 10) * 10, 1);
                                    }
                                    else
                                    {
                                        bc[(((int)ary[Index1] + 1) / 10) * 10] += 1;
                                    }
                                    /*
                                    if (gc[(int)ary[Index2]] == 1)
                                    {
                                        gc.Remove((int)ary[Index2]);
                                    }
                                    else
                                    {
                                        gc[(int)ary[Index2]]--;
                                    }
                                    */
                                    gc[(int)ary[Index2]]--;
                                    //ary.Remove((int)ary[Index2]);

                                    if (!bc.ContainsKey(((int)ary[Index2] + 1) / 10 * 10))
                                    {
                                        bc.Add((((int)ary[Index2]) / 10) * 10, 1);
                                    }
                                    else
                                    {
                                        bc[(((int)ary[Index2] + 1) / 10) * 10] += 1;
                                    }
                                    /*
                                    if (gc[(int)ary[Index3]] == 1)
                                    {
                                        gc.Remove((int)ary[Index3]);
                                    }
                                    else
                                    {
                                        gc[(int)ary[Index3]]--;
                                    }
                                    */
                                    gc[(int)ary[Index3]]--;
                                    //ary.Remove((int)ary[Index3]);

                                    if (!bc.ContainsKey(((int)ary[Index3] + 1) / 10 * 10))
                                    {
                                        bc.Add((((int)ary[Index3]) / 10) * 10, 1);
                                    }
                                    else
                                    {
                                        bc[(((int)ary[Index3] + 1) / 10) * 10] += 1;
                                    }
                                    /*
                                    if (gc[(int)ary[Index4]] == 1)
                                    {
                                        gc.Remove((int)ary[Index4]);
                                    }
                                    else
                                    {
                                        gc[(int)ary[Index4]]--;
                                    }
                                    */
                                    gc[(int)ary[Index4]]--;
                                    //ary.Remove((int)ary[Index4]);

                                    if (!bc.ContainsKey(((int)ary[Index4] + 1) / 10 * 10))
                                    {
                                        bc.Add((((int)ary[Index4]) / 10) * 10, 1);
                                    }
                                    else
                                    {
                                        bc[(((int)ary[Index4] + 1) / 10) * 10] += 1;
                                    }
                                    leftlen = 0;
                                    string tr = "";
                                    tr = ary[Index1] + "," + ary[Index2] + "," + ary[Index3] + "," + ary[Index4];

                                    string[] trs = tr.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                                    foreach (string trk in trs)
                                    {
                                        ary.Remove(int.Parse(trk));
                                    }
                                    break;
                                }
                            }
                        }
                    }
                }
            }

            //5块能解决
            if (gd == false && ary.Count >= 4 && leftlen > 0)
            {
                for (int Index1 = 0; Index1 <= ary.Count - 5 && gd == false; Index1++)
                {
                    for (int Index2 = 1; Index2 <= ary.Count - 4 && gd == false; Index2++)
                    {
                        for (int Index3 = 2; Index3 <= ary.Count - 3 && gd == false; Index3++)
                        {
                            for (int Index4 = 3; Index4 <= ary.Count - 2 && gd == false; Index4++)
                            {
                                for (int Index5 = 4; Index5 <= ary.Count - 1 && gd == false; Index5++)
                                {
                                    if ((int)ary[Index1] + (int)ary[Index2] + (int)ary[Index3] + (int)ary[Index4] + (int)ary[Index5] == leftlen)
                                    {
                                        gd = true;
                                        choose.Add(ary[Index1]);
                                        choose.Add(ary[Index2]);
                                        choose.Add(ary[Index3]);
                                        choose.Add(ary[Index4]);
                                        choose.Add(ary[Index5]);
                                        /*
                                        if (gc[(int)ary[Index1]] == 1)
                                        {
                                            gc.Remove((int)ary[Index1]);
                                        }
                                        else
                                        {
                                            gc[(int)ary[Index1]]--;
                                        }
                                        */
                                        gc[(int)ary[Index1]]--;
                                        //ary.Remove((int)ary[Index1]);

                                        if (!bc.ContainsKey(((int)ary[Index1] + 1) / 10 * 10))
                                        {
                                            bc.Add((((int)ary[Index1]) / 10) * 10, 1);
                                        }
                                        else
                                        {
                                            bc[(((int)ary[Index1] + 1) / 10) * 10] += 1;
                                        }
                                        /*
                                        if (gc[(int)ary[Index2]] == 1)
                                        {
                                            gc.Remove((int)ary[Index2]);
                                        }
                                        else
                                        {
                                            gc[(int)ary[Index2]]--;
                                        }
                                        */
                                        gc[(int)ary[Index2]]--;
                                        //ary.Remove((int)ary[Index2]);

                                        if (!bc.ContainsKey(((int)ary[Index2] + 1) / 10 * 10))
                                        {
                                            bc.Add((((int)ary[Index2]) / 10) * 10, 1);
                                        }
                                        else
                                        {
                                            bc[(((int)ary[Index2] + 1) / 10) * 10] += 1;
                                        }
                                        /*
                                        if (gc[(int)ary[Index3]] == 1)
                                        {
                                            gc.Remove((int)ary[Index3]);
                                        }
                                        else
                                        {
                                            gc[(int)ary[Index3]]--;
                                        }
                                        */
                                        gc[(int)ary[Index3]]--;
                                        //ary.Remove((int)ary[Index3]);

                                        if (!bc.ContainsKey(((int)ary[Index3] + 1) / 10 * 10))
                                        {
                                            bc.Add((((int)ary[Index3]) / 10) * 10, 1);
                                        }
                                        else
                                        {
                                            bc[(((int)ary[Index3] + 1) / 10) * 10] += 1;
                                        }
                                        /*
                                        if (gc[(int)ary[Index4]] == 1)
                                        {
                                            gc.Remove((int)ary[Index4]);
                                        }
                                        else
                                        {
                                            gc[(int)ary[Index4]]--;
                                        }
                                        */
                                        gc[(int)ary[Index4]]--;
                                        //ary.Remove((int)ary[Index4]);

                                        if (!bc.ContainsKey(((int)ary[Index4] + 1) / 10 * 10))
                                        {
                                            bc.Add((((int)ary[Index4]) / 10) * 10, 1);
                                        }
                                        else
                                        {
                                            bc[(((int)ary[Index4] + 1) / 10) * 10] += 1;
                                        }
                                        /*
                                        if (gc[(int)ary[Index5]] == 1)
                                        {
                                            gc.Remove((int)ary[Index5]);
                                        }
                                        else
                                        {
                                            gc[(int)ary[Index5]]--;
                                        }
                                        */
                                        gc[(int)ary[Index5]]--;
                                        //ary.Remove((int)ary[Index5]);

                                        if (!bc.ContainsKey(((int)ary[Index5] + 1) / 10 * 10))
                                        {
                                            bc.Add((((int)ary[Index5]) / 10) * 10, 1);
                                        }
                                        else
                                        {
                                            bc[(((int)ary[Index5] + 1) / 10) * 10] += 1;
                                        }
                                        leftlen = 0;
                                        string tr = "";
                                        tr = ary[Index1] + "," + ary[Index2] + "," + ary[Index3] + "," + ary[Index4] + "," + ary[Index5];

                                        string[] trs = tr.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                                        foreach (string trk in trs)
                                        {
                                            ary.Remove(int.Parse(trk));
                                        }
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
            }


            keyNames = "";

            foreach (int key in gc.Keys)

            {

                if (gc[key] == 0)

                {
                    keyNames += key + ",";
                }

            }

            str_keyNames = keyNames.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (string key in str_keyNames)
            {
                gc.Remove(int.Parse(key));
            }

            ary.Sort();
            ary.Reverse();


            keyNames = "";

            foreach (int key in gc.Keys)

            {

                if (gc[key] == 0)

                {
                    keyNames += key + ",";
                }

            }

            str_keyNames = keyNames.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (string key in str_keyNames)
            {
                gc.Remove(int.Parse(key));
            }
            //没有正合适的组合
            //剩的长大
            //优先选50的
            while (ary.Count > 0 && leftlen - (int)ary[ary.Count - 1] > 199)
            {
                for (int i = 0; i < ary.Count; i++)
                {
                    if (leftlen - (int)ary[i] > 199 && (leftlen - (int)ary[i]) % 100 < 1)
                    {
                        choose.Add(ary[i]);

                        leftlen -= (int)ary[i];

                        gc.Remove((int)ary[i]);

                        if (!bc.ContainsKey(((int)ary[i] + 1) / 10 * 10))
                        {
                            bc.Add((((int)ary[i]) / 10) * 10, 1);
                        }
                        else
                        {
                            bc[(((int)ary[i] + 1) / 10) * 10] += 1;
                        }
                        ary.RemoveAt(i);
                    }
                }
                for (int i = 0; i < ary.Count; i++)
                {
                    if (leftlen - (int)ary[i] > 199)
                    {
                        choose.Add(ary[i]);

                        leftlen -= (int)ary[i];
                        gc.Remove((int)ary[i]);

                        if (!bc.ContainsKey(((int)ary[i] + 1) / 10 * 10))
                        {
                            bc.Add((((int)ary[i]) / 10) * 10, 1);
                        }
                        else
                        {
                            bc[(((int)ary[i] + 1) / 10) * 10] += 1;
                        }
                        ary.RemoveAt(i);
                    }
                }
            }
            //剩的长小
            for (int i = 0; i < ary.Count && ary.Count > 0 && leftlen > 0; i++)
            {
                if ((int)ary[i] - leftlen > 199 && ((int)ary[i] - leftlen) % 100 < 1)
                {

                    choose.Add(leftlen);
                    if (gc[(int)ary[i]]-- == 0)
                    {
                        gc.Remove((int)ary[i]);
                    }
                    if (!gc.ContainsKey((int)ary[i] - leftlen))
                    {
                        gc.Add((int)ary[i] - leftlen, 1);
                    }
                    else
                    {
                        gc[(int)ary[i] - leftlen] += 1;

                    }

                    if (!bc.ContainsKey(leftlen))
                    {
                        bc.Add(leftlen, 1);
                    }
                    else
                    {
                        bc[leftlen] += 1;
                    }
                    ary[i] = (int)ary[i] - leftlen;
                    leftlen = 0;

                }
            }

            for (int i = 0; i < ary.Count && ary.Count > 0 && leftlen > 0; i++)
            {
                if ((int)ary[i] - leftlen > 199)
                {

                    choose.Add(leftlen);
                    if (gc[(int)ary[i]]-- == 0)
                    {
                        gc.Remove((int)ary[i]);
                    }
                    if (!gc.ContainsKey((int)ary[i] - leftlen))
                    {
                        gc.Add((int)ary[i] - leftlen, 1);
                    }
                    else
                    {
                        gc[(int)ary[i] - leftlen] += 1;

                    }

                    if (!bc.ContainsKey(leftlen))
                    {
                        bc.Add(leftlen, 1);
                    }
                    else
                    {
                        bc[leftlen] += 1;
                    }
                    ary[i] = (int)ary[i] - leftlen;
                    leftlen = 0;

                }
            }

            for (int i = 0; i < ary.Count && ary.Count > 0 && leftlen > 0; i++)
            {
                if (Math.Abs((int)ary[i] - leftlen) < 3)
                {

                    choose.Add(leftlen);
                    if (gc[(int)ary[i]]-- == 0)
                    {
                        gc.Remove((int)ary[i]);
                    }
                    if (!gc.ContainsKey((int)ary[i] - leftlen))
                    {
                        gc.Add((int)ary[i] - leftlen, 1);
                    }
                    else
                    {
                        gc[(int)ary[i] - leftlen] += 1;

                    }

                    if (!bc.ContainsKey(leftlen))
                    {
                        bc.Add(leftlen, 1);
                    }
                    else
                    {
                        bc[leftlen] += 1;
                    }
                    ary[i] = (int)ary[i] - leftlen;
                    leftlen = 0;

                }
            }
            //割剩
            if (leftlen != 0)
            {
                ary.Add(1200 - leftlen);
                choose.Add(leftlen);

                if (!gc.ContainsKey(1200 - leftlen))
                {
                    gc.Add(1200 - leftlen, 1);
                }
                else
                {
                    gc[1200 - leftlen] += 1;

                }

                if (!bc.ContainsKey(leftlen))
                {
                    bc.Add(leftlen, 1);
                }
                else
                {
                    bc[leftlen] += 1;
                }
                leftlen = 0;

            }


            keyNames = "";
            foreach (int key in gc.Keys)

            {
                if (gc[key] == 0)

                {
                    keyNames += key + ",";
                }

            }

            str_keyNames = keyNames.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (string key in str_keyNames)
            {
                gc.Remove(int.Parse(key));
            }


            //长板切小
            //表
            string qieS = "";
            foreach (int key in bc.Keys)

            {
                if (key > 601)

                {
                    qieS += key + ",";
                }

            }
            string[] qieSSZ = qieS.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string key in qieSSZ)
            {

                if (!bc.ContainsKey(int.Parse(key) / 100 / 2 * 100))
                {
                    bc.Add(int.Parse(key) / 100 / 2 * 100, bc[int.Parse(key)]);
                }
                else
                {
                    bc[int.Parse(key) / 100 / 2 * 100] += bc[int.Parse(key)];
                }

                if (!bc.ContainsKey(int.Parse(key) - int.Parse(key) / 100 / 2 * 100))
                {
                    bc.Add(int.Parse(key) - int.Parse(key) / 100 / 2 * 100, bc[int.Parse(key)]);
                }
                else
                {
                    bc[int.Parse(key) - int.Parse(key) / 100 / 2 * 100] += bc[int.Parse(key)];
                }

                bc.Remove(int.Parse(key));

            }

            //图
            int chcnt = choose.Count;
            ArrayList Torm = new ArrayList();
            for (int each = 0; each < chcnt; each++)
            {
                if ((int)choose[each] > 601)
                {
                    Torm.Add(choose[each]);
                }
            }
            foreach (int rm in Torm)
            {
                choose.Remove(rm);
                choose.Add(rm / 100 / 2 * 100);
                choose.Add(rm - rm / 100 / 2 * 100);
            }
            //画出来

            while (choose.Count > 0)
            {
                int line_section = (int)choose[0];


                if (Math.Abs(a.X - b.X) < 1)
                {//竖
                    if (a.Y > b.Y)//上向下
                    {
                        ori = a;
                        a = new Point2d(a.X, a.Y - line_section);
                    }
                    else//下向上
                    {
                        ori = a;
                        a = new Point2d(a.X, a.Y + line_section);
                    }

                }
                else//横
                {
                    if (a.X < b.X)//左向右
                    {
                        ori = a;
                        a = new Point2d(a.X + line_section, a.Y);
                    }
                    else//右向左
                    {
                        ori = a;
                        a = new Point2d(a.X - line_section, a.Y);
                    }
                }




                AlignedDimension acRotDime = new AlignedDimension();
                acRotDime.DimensionStyle = acCurDb.Dimstyle;
                acRotDime.Dimclrd = Autodesk.AutoCAD.Colors.Color.FromRgb(124, 252, 0);
                acRotDime.Dimclrt = Autodesk.AutoCAD.Colors.Color.FromRgb(124, 252, 0);
                acRotDime.Dimclre = Autodesk.AutoCAD.Colors.Color.FromRgb(124, 252, 0);


                acRotDime.XLine1Point = new Point3d(ori.X, ori.Y, 0);
                acRotDime.XLine2Point = new Point3d(a.X, a.Y, 0);


                if (Math.Abs(ori.X - a.X) < 1.0)//竖
                {
                    acRotDime.DimLinePoint = new Point3d(ori.X + textposition, (a.Y + ori.Y) / 2, 0);
                }
                else
                {
                    acRotDime.DimLinePoint = new Point3d((ori.X + a.X) / 2, a.Y + textposition, 0);
                }
                // 将新对象添加到块表记录 ModelSpace 及事务
                acBlkTblRec.AppendEntity(acRotDime);
                acTrans.AddNewlyCreatedDBObject(acRotDime, true);

                //finally
                //globalChoose.Add(line_section);
                choose.Remove(line_section);
            }



            ////加到材料切割数组里
            //if (!bc.ContainsKey((int)((int)a.GetDistanceTo(b) + 1) / 10 * 10))
            //{
            //    bc.Add(((int)((int)a.GetDistanceTo(b) + 1) / 10) * 10, 1);
            //}
            //else
            //{
            //    bc[((int)((int)a.GetDistanceTo(b) + 1) / 10) * 10] += 1;
            //}
            ////画出来
            //if (Math.Abs(a.X - b.X) < 1)
            //{//竖
            //    if (a.Y > b.Y)//上向下
            //    {
            //        ori = a;
            //        a = new Point2d(a.X, a.Y - a.GetDistanceTo(b));
            //    }
            //    else//下向上
            //    {
            //        ori = a;
            //        a = new Point2d(a.X, a.Y + a.GetDistanceTo(b) );
            //    }

            //}
            //else//横
            //{
            //    if (a.X < b.X)//左向右
            //    {
            //        ori = a;
            //        a = new Point2d(a.X + a.GetDistanceTo(b), a.Y );
            //    }
            //    else//右向左
            //    {
            //        ori = a;
            //        a = new Point2d(a.X - a.GetDistanceTo(b), a.Y );
            //    }
            //}




            //AlignedDimension acRotDime = new AlignedDimension();
            //acRotDime.DimensionStyle = acCurDb.Dimstyle;
            //acRotDime.Dimclrd = Autodesk.AutoCAD.Colors.Color.FromRgb(124, 252, 0);
            //acRotDime.Dimclrt = Autodesk.AutoCAD.Colors.Color.FromRgb(124, 252, 0);
            //acRotDime.Dimclre = Autodesk.AutoCAD.Colors.Color.FromRgb(124, 252, 0);


            //acRotDime.XLine1Point = new Point3d(ori.X, ori.Y, 0);
            //acRotDime.XLine2Point = new Point3d(a.X, a.Y, 0);


            //if (Math.Abs(ori.X - a.X) < 1.0)//竖
            //{
            //    acRotDime.DimLinePoint = new Point3d(ori.X + textposition, (a.Y + ori.Y) / 2, 0);
            //}
            //else
            //{
            //    acRotDime.DimLinePoint = new Point3d((ori.X + a.X) / 2, a.Y + textposition, 0);
            //}
            //// 将新对象添加到块表记录 ModelSpace 及事务
            //acBlkTblRec.AppendEntity(acRotDime);
            //acTrans.AddNewlyCreatedDBObject(acRotDime, true);
            //bc[(a.GetDistanceTo(b) + 1) % 10 * 10] += 1;


        }

        [CommandMethod("ListEntities")]
        public static void ListEntities()
        {


            // Get the current document and database, and start a transaction
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // 以读模式打开块表
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                // 以写模式打开块表记录 ModelSpace（模型空间）
                BlockTableRecord acBlkTblRec;
                acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                OpenMode.ForWrite) as BlockTableRecord;
                int nCnt = 0;
                acDoc.Editor.WriteMessage("\nModel space objects: ");
                bool minflag;
                // Step through each object in Model space and
                // display the type of object found
                foreach (ObjectId asObjId in acBlkTblRec)
                {
                    //asObjId.ToString
                    if (asObjId.ObjectClass.DxfName.Equals("HATCH"))  //LWPOLYLINE多段线  HATCH填充
                    {
                        //acDoc.Editor.WriteMessage("\nDXF name: " +asObjId.ObjectClass.DxfName);  
                        //acDoc.Editor.WriteMessage("\nObjectID: " + asObjId.ToString());
                        //acDoc.Editor.WriteMessage("\nHandle: " + asObjId.Handle.ToString());
                        //acDoc.Editor.WriteMessage("\n");
                        minflag = false;
                        Hatch hatch = asObjId.GetObject(OpenMode.ForRead) as Hatch;
                        HatchObjectType type = hatch.HatchObjectType;  //HatchObjectType: 
                        HatchStyle style = hatch.HatchStyle;  //HatchObjecthatch.HatchStyle: Normal
                        int lines = hatch.NumberOfHatchLines; //hatch.NumberOfHatchLines: 3
                        string stylename = hatch.PlotStyleName;   //hatch.PlotStyleName: ByBlock
                        //hatch.PatternName
                        if (hatch.PatternName.Equals("ANSI37") || hatch.PatternName.Equals("多孔材料"))
                        {
                            //                        acDoc.Editor.WriteMessage("\n HatchObjectType: " + hatch.HatchObjectType + " hatch.HatchStyle: " + hatch.HatchStyle
                            //+ " hatch.NumberOfHatchLines: " + hatch.NumberOfHatchLines + " hatch.PlotStyleName: " + hatch.PlotStyleName +
                            //" hatch.BlockName: " + hatch.BlockName + " hatch.PatternName:" + hatch.PatternName   + " hatch.Bounds" + hatch.Bounds + " hatch.NumberOfLoops:" + hatch.NumberOfLoops + "\n");
                            //hatch数量

                            //Line2dCollection lineCollection= hatch.GetHatchLinesData();

                            hatch.Highlight();
                            //获取hatch边界
                            int loopNum = hatch.NumberOfLoops;
                            acDoc.Editor.WriteMessage("\nloopNum: " + loopNum);
                            Point2dCollection col_point2d = new Point2dCollection();
                            BulgeVertexCollection col_ver = new BulgeVertexCollection();
                            Curve2dCollection col_cur2d = new Curve2dCollection();
                            Point2dCollection longest_point2d = new Point2dCollection();
                            Point2dCollection eXchangeTmp = new Point2dCollection();
                            Point2dCollection drawLinePoints = new Point2dCollection();
                            float[] textPosition = new float[0xaa];

                            double longestline = 0;
                            for (int i = 0; i < loopNum; i++)
                            {
                                col_point2d.Clear();
                                HatchLoop hatLoop = null;
                                try
                                {
                                    hatLoop = hatch.GetLoopAt(i);
                                    acDoc.Editor.WriteMessage("\n" + hatLoop.Curves.ToString());
                                }
                                catch (System.Exception)
                                {
                                    continue;

                                }

                                //如果HatchLoop为PolyLine  
                                if (hatLoop.IsPolyline)
                                {
                                    acDoc.Editor.WriteMessage("\nPolyline style");
                                    col_ver = hatLoop.Polyline;
                                    foreach (BulgeVertex vertex in col_ver)
                                    {
                                        col_point2d.Add(vertex.Vertex);
                                    }
                                }
                                //如果HatchLoop为Curves  
                                else
                                {
                                    col_cur2d = hatLoop.Curves;
                                    acDoc.Editor.WriteMessage("\n hatLoop.Curves.Count:" + hatLoop.Curves.Count);
                                    foreach (Curve2d item in col_cur2d)
                                    {
                                        Point2d[] M_point2d = item.GetSamplePoints(2);
                                        //item.StartPoint
                                        foreach (Point2d point in M_point2d)
                                        {
                                            if (!col_point2d.Contains(point))
                                                col_point2d.Add(point);
                                            //if(hatLoop.Curves.Count>8)

                                            //{
                                            //    acDoc.Editor.WriteMessage("\n point:" + point);
                                            // acDoc.Editor.WriteMessage("\n start point:" + item.StartPoint + "\n end point:" + item.EndPoint);
                                            //}
                                        }

                                    }
                                    //acDoc.Editor.WriteMessage("\nNot Polyline style, col_point2d.Count: "+ col_point2d.Count); 
                                }
                                //去杂点
                                int pindex = 0, ppindex = 1;
                                for (pindex = 0; pindex < col_point2d.Count; pindex++)
                                {
                                    for (ppindex = pindex + 1; ppindex < col_point2d.Count; ppindex++)
                                    {
                                        if (col_point2d[pindex].GetDistanceTo(col_point2d[ppindex]) < 1)
                                        {
                                            col_point2d.Remove(col_point2d[ppindex]);
                                        }
                                    }

                                    //acDoc.Editor.WriteMessage("\n col_point2d: " + col_point2d[pindex]);
                                    //col_point2d[pindex].GetDistanceTo(col_point2d[(pindex + 1) % col_point2d.Count]);
                                }
                                //去非顶点的点
                                pindex = 0;
                                ppindex = 1;
                                int pppindex = 2;
                                ArrayList p2rm = new ArrayList();
                                while ( pindex < col_point2d.Count - 2)
                                {

                                    if ((Math.Abs(col_point2d[pindex].X - col_point2d[ppindex].X) < 1 && Math.Abs(col_point2d[ppindex].X - col_point2d[pppindex].X) < 1)||( Math.Abs(col_point2d[pindex].Y - col_point2d[ppindex].Y) < 1 && Math.Abs(col_point2d[ppindex].Y - col_point2d[pppindex].Y) < 1))
                                    {
                                        p2rm.Add(col_point2d[ppindex]);
                                        acDoc.Editor.WriteMessage("\n pindex: " + pindex + "\n ppindex: " + ppindex + "\n pppindex: " + pppindex);
                                        ppindex++;
                                        pppindex++;
                                    }
                                    else
                                    {
                                        pppindex++;
                                        ppindex++;
                                        pindex++;
                                    }


                                    //acDoc.Editor.WriteMessage("\n col_point2d: " + col_point2d[pindex]);
                                    //col_point2d[pindex].GetDistanceTo(col_point2d[(pindex + 1) % col_point2d.Count]);
                                }
                                foreach(Point2d trm in p2rm)
                                {
                                    acDoc.Editor.WriteMessage("\n trm: " + trm);
                                    if(col_point2d.Contains(trm))
                                    {
                                        col_point2d.Remove(trm);
                                    }
                                }


                                if (col_point2d.Count % 2 == 1)
                                {
                                    break;
                                }

                                //将线长按点的顺序放入patchline 数组
                                double[] hatchlines = new double[col_point2d.Count];
                                for (pindex = 0; pindex < col_point2d.Count; pindex++)
                                {
                                    hatchlines[pindex] = col_point2d[pindex].GetDistanceTo(col_point2d[(pindex + 1) % col_point2d.Count]);

                                }
                                acDoc.Editor.WriteMessage("\n  col_point2d.Count: " + col_point2d.Count);
                                //======================================================================================================
                                //double maxValue = patchlines[0];
                                double minValue = 200;
                                int minIndex = 0;
                                double totalLen = 0;
                                int width1Index = 0;
                                int width2Index = 0;
                                for (int ia = 0; ia < hatchlines.Length; ia++)
                                {
                                    //if (patchlines[ia] > maxValue) maxValue = patchlines[ia];
                                    //找最小边长（宽）
                                    //if ((hatchlines[ia] < minValue) )
                                    if ((hatchlines[ia] < minValue) && (hatchlines[ia] > 59) && (hatchlines[ia] < 151) && Math.Abs(hatchlines[ia] - hatchlines[(ia + hatchlines.Length / 2) % hatchlines.Length]) < 1)
                                    {
                                        minValue = hatchlines[ia];
                                        minIndex = ia;
                                        minflag = true;
                                    }
                                    width1Index = ia;
                                    width2Index = (ia + hatchlines.Length / 2) % hatchlines.Length;
                                    int widthtmp;
                                    if (width1Index < width2Index)
                                    {
                                        widthtmp = width1Index;
                                        width1Index = width2Index;
                                        width2Index = width1Index;

                                    }
                                    totalLen += hatchlines[ia];
                                    acDoc.Editor.WriteMessage("\n hatchlines[ia] " + hatchlines[ia]);

                                }
                                acDoc.Editor.WriteMessage("\n  minValue: " + minValue);
                                if (minflag == false)
                                {
                                    //break;
                                }
                                /*
                                if (longestline < totalLen)
                                {
                                    eXchangeTmp = longest_point2d;
                                    longest_point2d = col_point2d;
                                    longestline = totalLen;
                                    col_point2d = eXchangeTmp;
                                    if (eXchangeTmp.Count == 0)
                                    {
                                        break;
                                    }
                                }
                                */

                                acDoc.Editor.WriteMessage("\n width1Index " + width1Index + " width2Index " + width2Index);

                                if (totalLen < 400.0 + 1.0)
                                {
                                    break;
                                }

                                nCnt = nCnt + 1;


 



                                totalLen = totalLen / 2 - minValue;
                                acDoc.Editor.WriteMessage("\n totalLen: " + totalLen);
                                double hatchWidth = minValue;
                                int hatchWidthIndex = minIndex;
                                acDoc.Editor.WriteMessage("\n col_point2d.Count " + col_point2d.Count + " hatchWidthIndex  " + hatchWidthIndex);

                                //依次处理两边的边,过程中边长会变化。
                                int pline = hatchlines.Length / 2 - 1;//一侧有几条边n
                                for (int pcount = 1; pcount < pline + 1; pcount++)//起始点和结束点固定
                                {

                                    int thisIndexL = (minIndex - pcount + col_point2d.Count) % col_point2d.Count;//index递减
                                    int thisIndexR = (minIndex + pcount) % col_point2d.Count;//index递增

                                    //选出本次长边

                                    if (Math.Abs(hatchlines[thisIndexL] - hatchlines[thisIndexR]) < 1.0)//两边相等，结束；
                                    {
                                        acDoc.Editor.WriteMessage("\n 结束");
                                        int endIndex = thisIndexL > thisIndexR ? thisIndexL : thisIndexR;
                                        drawLinePoints.Add(col_point2d[thisIndexR]);
                                        drawLinePoints.Add(col_point2d[(thisIndexR + 1) % col_point2d.Count]);
                                        break;
                                    }


                                    int nextIndexL = (minIndex - pcount - 1 + col_point2d.Count) % col_point2d.Count;
                                    int nextIndexR = (minIndex + pcount + 1) % col_point2d.Count;
                                    acDoc.Editor.WriteMessage("\n thisIndexL: " + thisIndexL + " thisIndexR  " + thisIndexR);
                                    acDoc.Editor.WriteMessage("\n nextIndexL: " + nextIndexL + " nextIndexR  " + nextIndexR);

                                    double longLen = hatchlines[thisIndexL] > hatchlines[thisIndexR] ? hatchlines[thisIndexL] : hatchlines[thisIndexR];
                                    int longLenIndex = hatchlines[thisIndexL] > hatchlines[thisIndexR] ? thisIndexL : thisIndexR;

                                    int otherIndex = longLenIndex == thisIndexL ? thisIndexR : thisIndexL;
                                    int step = minIndex - longLenIndex;
                                    acDoc.Editor.WriteMessage("\n longLenIndex: " + longLenIndex);
                                    int nextIndex;
                                    int nnIndex;
                                    int nextOtherIndex;
                                    int nnOtherIndex;
                                    if (longLenIndex == thisIndexL)
                                    //if (longLenIndex - minIndex < 0 || longLenIndex - minIndex == col_cur2d.Count - 1)//编号小的，编号点离的远
                                    {
                                        nextIndex = longLenIndex;
                                        nnIndex = (longLenIndex - 1 + col_point2d.Count) % col_point2d.Count;
                                        nextOtherIndex = (otherIndex + 1) % col_point2d.Count;
                                        nnOtherIndex = (otherIndex + 2) % col_point2d.Count;
                                        longLenIndex = (longLenIndex + 1) % col_point2d.Count;
                                        drawLinePoints.Add(new Point2d(col_point2d[(longLenIndex) % col_point2d.Count].X, col_point2d[(longLenIndex) % col_point2d.Count].Y));

                                    }
                                    //if (longLenIndex - minIndex >0 || longLenIndex - minIndex == 1 - col_cur2d.Count)//编号大，编号+1的点变化
                                    else
                                    {
                                        nextIndex = (longLenIndex + 1) % col_point2d.Count;
                                        nnIndex = (longLenIndex + 2) % col_point2d.Count;
                                        nextOtherIndex = otherIndex;
                                        nnOtherIndex = (otherIndex - 1 + col_point2d.Count) % col_point2d.Count;
                                        drawLinePoints.Add(new Point2d(col_point2d[(longLenIndex) % col_point2d.Count].X, col_point2d[(longLenIndex) % col_point2d.Count].Y));

                                    }
                                    //double nextline;

                                    //选出下次长边,没用 选对边----------
                                    bool nextEqualFlag = false;
                                    double longLenNext = hatchlines[nextIndexL] > hatchlines[nextIndexR] ? hatchlines[nextIndexL] : hatchlines[nextIndexR];
                                    int longLenIndexNext = hatchlines[nextIndexL] > hatchlines[nextIndexR] ? nextIndexL : nextIndexR;
                                    acDoc.Editor.WriteMessage("\n longLenIndexNext: " + longLenIndexNext);

                                    //if (Math.Abs(hatchlines[nextIndexL] - hatchlines[nextIndexR]) < 1.0)
                                    //{
                                    //    nextEqualFlag = true;
                                    //    longLenNext += width1Index;

                                    //}
                                    //acDoc.Editor.WriteMessage("\n 两次边长: " + longLen + "   " + longLenNext);
                                    ////比同边下一条直线长度,顶点序号可能顺时针可能逆时针 
                                    //if (longLenIndex < width1Index && longLenIndex > width2Index)//同向
                                    //{
                                    //    nextIndex = (longLenIndex + 1 + col_point2d.Count) % col_point2d.Count;
                                    //    nnIndex = (longLenIndex + 2 + col_point2d.Count) % col_point2d.Count;
                                    //    nextOtherIndex = (otherIndex - 1) % col_point2d.Count;
                                    //    nnOtherIndex = (otherIndex - 2) % col_point2d.Count;
                                    //    //longLenNext = hatchlines[nextIndex];
                                    //}
                                    //else//反向
                                    //{
                                    //    nextIndex = (longLenIndex - 1) % col_point2d.Count;
                                    //    nnIndex = (longLenIndex - 2) % col_point2d.Count;
                                    //    nextOtherIndex = (otherIndex + 1 + col_point2d.Count) % col_point2d.Count;
                                    //    nnOtherIndex = (otherIndex + 2 + col_point2d.Count) % col_point2d.Count;
                                    //    //snextline = hatchlines[nextIndex];
                                    //}
                                    //处理长度
                                    acDoc.Editor.WriteMessage("\n 开始处理长度 ");
                                    acDoc.Editor.WriteMessage("\n longLenIndex " + longLenIndex + "  nextIndex :" + nextIndex + " nnIndex: " + nnIndex + " nextOtherIndex: " + nextOtherIndex + " nnotherIndex: " + nnOtherIndex);

                                    if (longLen > longLenNext + 1.0)//下条线短  ，长边减w加到draw，下条对边加
                                    {
                                        //判断线向
                                        acDoc.Editor.WriteMessage("\n 判断线向 ");
                                        //acDoc.Editor.WriteMessage("\n col_point2d[longLenIndex] " + col_point2d[longLenIndex] + " col_point2d[(longLenIndex + 1) % col_point2d.Count] " + col_point2d[(longLenIndex + 1) % col_point2d.Count]);
                                        //if (Math.Abs(col_point2d[nextIndex].X - drawLinePoints[(pcount * 2 - 2)].X) < 1.0)//竖
                                        if (Math.Abs(col_point2d[nextIndex].X - col_point2d[(longLenIndex)].X) < 1.0)//竖

                                        {
                                            //本条边
                                            acDoc.Editor.WriteMessage("\n 竖 ");

                                            if (col_point2d[(longLenIndex)].Y > col_point2d[nextIndex].Y)//上向下
                                            {
                                                drawLinePoints.Add(new Point2d(col_point2d[nextIndex].X, col_point2d[nextIndex].Y + hatchWidth));

                                            }
                                            else//下向上
                                            {
                                                drawLinePoints.Add(new Point2d(col_point2d[nextIndex].X, col_point2d[nextIndex].Y - hatchWidth));
                                            }

                                            //设置标注偏移位置
                                            if (col_point2d[longLenIndex].X < col_point2d[otherIndex].X)
                                            {
                                                textPosition[pcount] = -200;
                                            }
                                            else
                                            {
                                                textPosition[pcount] = 200;
                                            }

                                        }
                                        else//横
                                        {
                                            acDoc.Editor.WriteMessage("\n 横");

                                            if (col_point2d[(longLenIndex)].X - col_point2d[nextIndex].X > 0)//右向左
                                            {
                                                drawLinePoints.Add(new Point2d(col_point2d[nextIndex].X + hatchWidth, col_point2d[nextIndex].Y));
                                            }
                                            else//左向右
                                            {
                                                drawLinePoints.Add(new Point2d(col_point2d[nextIndex].X - hatchWidth, col_point2d[nextIndex].Y));
                                            }
                                            //设置标注偏移位置
                                            if (col_point2d[longLenIndex].Y < col_point2d[otherIndex].Y)
                                            {
                                                textPosition[pcount] = +200;
                                            }
                                            else
                                            {
                                                textPosition[pcount] = -200;
                                            }
                                        }

                                        //n0-width n+1 0+width

                                        //--------------------------------------------------------------
                                        //下条对边加
                                        acDoc.Editor.WriteMessage("\n 下条对边加");

                                        if (Math.Abs(col_point2d[nextOtherIndex].X - col_point2d[nnOtherIndex].X) < 1.0)//竖
                                        {
                                            if (col_point2d[nextOtherIndex].Y - col_point2d[nnOtherIndex].Y > 0)//上向下
                                            {
                                                col_point2d[nextOtherIndex] = new Point2d(col_point2d[nextOtherIndex].X, col_point2d[nextOtherIndex].Y + hatchWidth);
                                            }
                                            else//下向上
                                            {
                                                col_point2d[nextOtherIndex] = new Point2d(col_point2d[nextOtherIndex].X, col_point2d[nextOtherIndex].Y - hatchWidth);
                                            }
                                            //col_point2d[nextOtherIndex] = new Point2d(col_point2d[nextOtherIndex].X, col_point2d[nextIndex].Y);
                                        }
                                        else//横
                                        {
                                            if (col_point2d[nextOtherIndex].X - col_point2d[nnOtherIndex].X > 0)//右向左
                                            {
                                                col_point2d[nextOtherIndex] = new Point2d(col_point2d[nextOtherIndex].X + hatchWidth, col_point2d[nextOtherIndex].Y);
                                            }
                                            else//左向右
                                            {
                                                col_point2d[nextOtherIndex] = new Point2d(col_point2d[nextOtherIndex].X - hatchWidth, col_point2d[nextOtherIndex].Y);
                                            }
                                            //col_point2d[nextOtherIndex] = new Point2d(col_point2d[nextIndex].X, col_point2d[nextOtherIndex].Y);
                                        }

                                    }
                                    else//下条线长，下条边减
                                    {
                                        //下条边减
                                        acDoc.Editor.WriteMessage("\n 下条边减");
                                        drawLinePoints.Add(new Point2d(col_point2d[nextIndex].X, col_point2d[nextIndex].Y));

                                        //判断线向
                                        //if (Math.Abs(col_point2d[longLenIndex].X - col_point2d[longLenIndex + 1].X) < 1.0)//竖
                                        //{
                                        //    acDoc.Editor.WriteMessage("\n 竖\n");

                                        //    //本条边
                                        //    if (col_point2d[otherIndex].Y - col_point2d[nextOtherIndex].Y > 0)//上向下
                                        //    {
                                        //        drawLinePoints[pcount * 2 - 1] = new Point2d(col_point2d[nextOtherIndex].X, col_point2d[nextOtherIndex].Y - hatchWidth);


                                        //    }
                                        //    else//下向上
                                        //    {
                                        //        drawLinePoints[pcount * 2 - 1] = new Point2d(col_point2d[nextOtherIndex].X, col_point2d[nextOtherIndex].Y + hatchWidth);
                                        //    }

                                        //}
                                        //else//横
                                        //{
                                        //    acDoc.Editor.WriteMessage("\n 横\n");

                                        //    if (col_point2d[otherIndex].X - col_point2d[nextOtherIndex].X > 0)//右向左
                                        //    {
                                        //        col_point2d[nextOtherIndex] = new Point2d(col_point2d[nextOtherIndex].X - hatchWidth, col_point2d[nextOtherIndex].Y);
                                        //    }
                                        //    else//左向右
                                        //    {
                                        //        col_point2d[nextOtherIndex] = new Point2d(col_point2d[nextOtherIndex].X + hatchWidth, col_point2d[nextOtherIndex].Y);
                                        //    }
                                        //    drawLinePoints[pcount * 2 - 2] = new Point2d(col_point2d[otherIndex].X, col_point2d[otherIndex].Y);

                                        //}

                                        //下条边减
                                        if (Math.Abs(col_point2d[nextIndex].X - col_point2d[nnIndex].X) < 1.0)//竖
                                        {
                                            if (col_point2d[nextIndex].Y - col_point2d[nnIndex].Y > 0)//上向下
                                            {
                                                col_point2d[nextIndex] = new Point2d(col_point2d[nextIndex].X, col_point2d[nextIndex].Y - hatchWidth);
                                            }
                                            else//下向上
                                            {
                                                col_point2d[nextIndex] = new Point2d(col_point2d[nextIndex].X, col_point2d[nextIndex].Y + hatchWidth);
                                            }
                                            //col_point2d[nextIndex] = new Point2d(col_point2d[nextOtherIndex].X, col_point2d[nextIndex].Y);
                                        }
                                        else//横
                                        {
                                            if (col_point2d[nextIndex].X - col_point2d[nnIndex].X > 0)//右向左
                                            {
                                                col_point2d[nextIndex] = new Point2d(col_point2d[nextIndex].X - hatchWidth, col_point2d[nextIndex].Y);
                                            }
                                            else//左向右
                                            {
                                                col_point2d[nextIndex] = new Point2d(col_point2d[nextIndex].X + hatchWidth, col_point2d[nextIndex].Y);
                                            }
                                            //col_point2d[nextIndex] = new Point2d(col_point2d[nextIndex].X, col_point2d[nextOtherIndex].Y);
                                        }
                                    }


                                    //更新边长数据
                                    for (pindex = 0; pindex < col_point2d.Count; pindex++)
                                    {
                                        hatchlines[pindex] = col_point2d[pindex].GetDistanceTo(col_point2d[(pindex + 1) % col_point2d.Count]);
                                    }
                                    acDoc.Editor.WriteMessage("\n  col_point2d.Count: " + col_point2d.Count);
                                    //-------------------------------------
                                    //acDoc.Editor.WriteMessage("\n  要标注的两点坐标： " + drawLinePoints[2 * linecount - 2] + "  " + drawLinePoints[2 * linecount - 21]);
                                    //标注：----------------------------------------------------


                                    //AlignedDimension acRotDim = new AlignedDimension();
                                    ////acRotDim.XLine2Point;
                                    //acRotDim.XLine2Point = new Point3d(drawLinePoints[0].X, drawLinePoints[0].Y, 0);
                                    //acRotDim.XLine2Point = new Point3d(drawLinePoints[1].X, drawLinePoints[1].Y, 0);
                                    //////acRotDim.Rotation = 0.707;
                                    //if (Math.Abs(drawLinePoints[0].X - drawLinePoints[0].X) < 1.0)
                                    //{
                                    //    acRotDim.DimLinePoint = new Point3d(0, drawLinePoints[0].X+2*hatchWidth, 0);
                                    //}
                                    //else
                                    //{
                                    //    acRotDim.DimLinePoint = new Point3d(0, drawLinePoints[0].Y + 2 * hatchWidth, 0);
                                    //}
                                    ////acRotDim.DimLinePoint = new Point3d(0, 5, 0);
                                    //acRotDim.DimensionStyle = acCurDb.Dimstyle;
                                    //// 将新对象添加到块表记录 ModelSpace 及事务
                                    //acBlkTblRec.AppendEntity(acRotDim);
                                    //acTrans.AddNewlyCreatedDBObject(acRotDim, true);
                                    // 提交修改，关闭事务
                                }
                                for (int linecount = 1; linecount < drawLinePoints.Count / 2 + 1; linecount++)
                                {
                                    //AlignedDimension acRotDim = new AlignedDimension();
                                    //acRotDim.Dimclrd = Autodesk.AutoCAD.Colors.Color.FromRgb(124, 252, 0);
                                    //acRotDim.Dimclrt = Autodesk.AutoCAD.Colors.Color.FromRgb(124, 252, 0);
                                    //acRotDim.Dimclre = Autodesk.AutoCAD.Colors.Color.FromRgb(124, 252, 0);
                                    ////acRotDim.Dimexe = 110;
                                    ////acRotDim.Dimexo = 100;
                                    ////acRotDim.XLine2Point
                                    //acRotDim.XLine1Point = new Point3d(drawLinePoints[2 * linecount - 2].X, drawLinePoints[2 * linecount - 2].Y, 0);
                                    //acRotDim.XLine2Point = new Point3d(drawLinePoints[2 * linecount - 1].X, drawLinePoints[2 * linecount - 1].Y, 0);
                                    ////acRotDim.Rotation = 0.707;
                                    //if (Math.Abs(drawLinePoints[2 * linecount - 2].X - drawLinePoints[2 * linecount - 1].X) < 1.0)//竖
                                    //{
                                    //    acRotDim.DimLinePoint = new Point3d(drawLinePoints[2 * linecount - 2 ].X + textPosition[linecount], (drawLinePoints[2 * linecount - 2].Y + drawLinePoints[2 * linecount - 1].Y) / 2, 0);
                                    //}
                                    //else
                                    //{
                                    //    acRotDim.DimLinePoint = new Point3d((drawLinePoints[2 * linecount - 1].X + drawLinePoints[2 * linecount - 2].X) / 2, drawLinePoints[2 * linecount - 1].Y + textPosition[linecount], 0);
                                    //}
                                    //acRotDim.DimensionStyle = acCurDb.Dimstyle;
                                    //// 将新对象添加到块表记录 ModelSpace 及事务
                                    //acBlkTblRec.AppendEntity(acRotDim);
                                    //acTrans.AddNewlyCreatedDBObject(acRotDim, true);












                                    // 提交修改，关闭事务
                                    //acTrans.Commit();
                                    if (drawLinePoints[2 * linecount - 2].GetDistanceTo(drawLinePoints[2 * linecount - 1]) < 80000)
                                    {
                                        cutLine(drawLinePoints[2 * linecount - 2], drawLinePoints[2 * linecount - 1], ref acCurDb, textPosition[linecount], acTrans, ref acBlkTblRec);
                                    }
                                }

                                acDoc.Editor.WriteMessage("\n find 依次处理两边的边 over ");

                                //alignpoint----------------------------------------------------
                                //for (int linecount = 1; linecount < col_point2d.Count / 2 - 1; linecount++)
                                //{
                                //    AlignedDimension acRotDim = new AlignedDimension();
                                //    //acRotDim.XLine2Point
                                //    acRotDim.XLine2Point = new Point3d(drawLinePoints[2 * linecount - 2].X, drawLinePoints[2 * linecount - 2].Y, 0);
                                //    acRotDim.XLine2Point = new Point3d(drawLinePoints[2 * linecount - 1].X, drawLinePoints[2 * linecount - 1].Y, 0);
                                //    //acRotDim.Rotation = 0.707;
                                //    //acRotDim.DimLinePoint = new Point3d(0, 5, 0);
                                //    acRotDim.DimensionStyle = acCurDb.Dimstyle;
                                //    // 将新对象添加到块表记录 ModelSpace 及事务
                                //    acBlkTblRec.AppendEntity(acRotDim);
                                //    acTrans.AddNewlyCreatedDBObject(acRotDim, true);
                                //    // 提交修改，关闭事务
                                //    acTrans.Commit();
                                //}----------------------------------------------------
                                //根据获得的Point2d点集创建闭合Polyline  
                                //Polyline pl = new Polyline();
                                //pl.Closed = true;
                                //pl.Color = hat.Color;
                                //PolylineTools.CreatePolyline(pl, col_point2d);
                                //btr.AppendEntity(pl);
                                //trans.AddNewlyCreatedDBObject(pl, true); 
                            }
                            if (minflag == false)
                            {
                                //continue;

                            }
                        }

                    }

                }

                bc.Remove(0);

                FileInfo newFile = new FileInfo(@"test.xlsx");
                if (newFile.Exists)
                {
                    newFile.Delete();
                    newFile = new FileInfo(@"test.xlsx");
                }

                using (ExcelPackage package = new ExcelPackage(newFile))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("test");
                    worksheet.Cells[1, 1].Value = "板材长度";
                    worksheet.Cells[1, 2].Value = "数量";
                    int n = 2;
                    foreach (KeyValuePair<int, int> kvp in bc)
                    {
                        //if(kvp.Key>601)
                        //{
                        //    worksheet.Cells[n, 1].Value = kvp.Key/2;
                        //    worksheet.Cells[n, 2].Value = kvp.Value*2;
                        //}
                        //else
                        {
                            worksheet.Cells[n, 1].Value = kvp.Key;
                            worksheet.Cells[n, 2].Value = kvp.Value;
                        }

                        n++;
                    }
                    worksheet.Cells[n, 1].Value = "废料长度";
                    worksheet.Cells[n++, 2].Value = "数量";
                    foreach (KeyValuePair<int, int> kvp in gc)
                    {

                        Console.WriteLine("废料长度：{0},数量：{1}", kvp.Key, kvp.Value);
                        worksheet.Cells[n, 1].Value = kvp.Key;
                        worksheet.Cells[n, 2].Value = kvp.Value;
                        n++;
                    }

                    package.Save();
                }
                acTrans.Commit();
                acDoc.Editor.WriteMessage("\n total hatch:" + nCnt);
                //图层
                //LayerTable acLayerTable = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead) as LayerTable;
                //foreach (ObjectId acObjId in acLayerTable)
                //{
                //    LayerTableRecord acLyrTblRec;
                //    acLyrTblRec = acTrans.GetObject(acObjId,
                //    OpenMode.ForRead) as LayerTableRecord;
                //    //acDoc.Editor.WriteMessage("\n" + acLyrTblRec.Name);
                //}
                ////ObjectId acObjId1 = acLayerTable["AE-FLOR"];
                ////acDoc.Editor.WriteMessage("\n" + acObjId1.ObjectClass.DxfName);
                ////Console.WriteLine("\n" + acObjId1.ObjectClass.DxfName);
                //if (acLayerTable.Has("AE-FLOR") == true)
                //{
                //    LayerTableRecord acLyrTblRec;
                //    acLyrTblRec = acTrans.GetObject(acLayerTable["AE-FLOR"],
                //    OpenMode.ForWrite) as
                //    LayerTableRecord;
                //    try
                //    {
                //        //acLyrTblRec.Erase();
                //        //acDoc.Editor.WriteMessage("\n'AE-FLOR' was erased");

                //        //    acDoc.Editor.WriteMessage("\n color :" +acLyrTblRec.EntityColor);
                //        //acDoc.Editor.WriteMessage("\n type:" + acLyrTblRec.GetType());
                //        //acDoc.Editor.WriteMessage("\n PlotStyleName:" + acLyrTblRec.PlotStyleName);
                //        //acDoc.Editor.WriteMessage("\n PlotStyleNameId:" + acLyrTblRec.PlotStyleNameId);

                //        // 提交修改
                //        acTrans.Commit();
                //    }
                //    catch
                //    {
                //        //acDoc.Editor.WriteMessage("\n'AE-FLOR' could not be erased");
                //    }
                //}
                //else
                //{
                //    //acDoc.Editor.WriteMessage("\n'AE-FLOR' does not exist");
                //}

                // If no objects are found then display a message
                if (nCnt == 0)
                {
                    acDoc.Editor.WriteMessage("\n  No objects found");
                }

                // Dispose of the transaction
            }
        }

        [CommandMethod("CalculateDefinedArea")]
        public static void CalculateDefinedArea()
        {
            // 提示用户输入 5 个点
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            PromptPointResult pPtRes;
            Point2dCollection colPt = new Point2dCollection();
            PromptPointOptions pPtOpts = new PromptPointOptions("");
            // 提示输入第 1 个点
            pPtOpts.Message = "\nSpecify first point: ";
            pPtRes = acDoc.Editor.GetPoint(pPtOpts);
            colPt.Add(new Point2d(pPtRes.Value.X, pPtRes.Value.Y));
            // 如果用户按 ESC 键或取消命令就退出
            if (pPtRes.Status == PromptStatus.Cancel) return;
            int nCounter = 1;
            while (nCounter <= 4)
            {
                // 提示下一个点
                switch (nCounter)
                {
                    case 1:
                        pPtOpts.Message = "\nSpecify second point: ";
                        break;
                    case 2:
                        pPtOpts.Message = "\nSpecify third point: ";
                        break;
                    case 3:
                        pPtOpts.Message = "\nSpecify fourth point: ";
                        break;
                    case 4:
                        pPtOpts.Message = "\nSpecify fifth point: ";
                        break;
                }
                // 用前一个点作基点
                pPtOpts.UseBasePoint = true;
                pPtOpts.BasePoint = pPtRes.Value;
                pPtRes = acDoc.Editor.GetPoint(pPtOpts);
                colPt.Add(new Point2d(pPtRes.Value.X, pPtRes.Value.Y));
                if (pPtRes.Status == PromptStatus.Cancel) return;
                // 计数加 1
                nCounter = nCounter + 1;
            }
            // 用 5 个点创建多段线
            // 所有的 2D 实体对象和 3D 实体对象都实现了 IDisposable，故可以使用 using 语句
            using (Polyline acPoly = new Polyline())
            {
                acPoly.AddVertexAt(0, colPt[0], 0, 0, 0);
                acPoly.AddVertexAt(1, colPt[1], 0, 0, 0);
                acPoly.AddVertexAt(2, colPt[2], 0, 0, 0);
                acPoly.AddVertexAt(3, colPt[3], 0, 0, 0);
                acPoly.AddVertexAt(4, colPt[4], 0, 0, 0);
                // 闭合多段线
                acPoly.Closed = true;
                // 查询多段线面积
                Application.ShowAlertDialog("Area of polyline: " +
                acPoly.Area.ToString());
                // 销毁多段线
            }
        }


        [CommandMethod("GetPointsFromUser")]
        public static void GetPointsFromUser()
        {
            // 获取当前数据库，启动事务管理器
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            PromptPointResult pPtRes;
            PromptPointOptions pPtOpts = new PromptPointOptions("");
            // 提示起点
            pPtOpts.Message = "\nEnter the start point of the line: ";
            pPtRes = acDoc.Editor.GetPoint(pPtOpts);
            Point3d ptStart = pPtRes.Value;
            // 如果用户按 ESC 键或取消命令，就退出
            if (pPtRes.Status == PromptStatus.Cancel) return;
            // 提示终点
            pPtOpts.Message = "\nEnter the end point of the line: ";
            pPtOpts.UseBasePoint = true;
            pPtOpts.BasePoint = ptStart;
            pPtRes = acDoc.Editor.GetPoint(pPtOpts);
            Point3d ptEnd = pPtRes.Value;
            if (pPtRes.Status == PromptStatus.Cancel) return;
            // 启动事务
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                BlockTable acBlkTbl;
                BlockTableRecord acBlkTblRec;
                // 以写模式打开模型空间
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                OpenMode.ForRead) as BlockTable;
                acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                OpenMode.ForWrite) as BlockTableRecord;
                // 创建直线
                Line acLine = new Line(ptStart, ptEnd);
                // 添加直线
                acBlkTblRec.AppendEntity(acLine);
                acTrans.AddNewlyCreatedDBObject(acLine, true);
                // 缩放图形到全部显示
                acDoc.SendStringToExecute("._zoom _all ", true, false, false);
                // 提交修改， 关闭事务
                acTrans.Commit();
            }
        }


        [CommandMethod("AddHatch")]
        public static void AddHatch()
        {
            // 获取当前文档和数据库
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            // 启动事务
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // 以读模式打开 Block 表
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                OpenMode.ForRead) as BlockTable;
                // 以写模式打开 Block 表记录 Model 空间
                BlockTableRecord acBlkTblRec;
                acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                OpenMode.ForWrite) as BlockTableRecord;
                // 创建一个圆作为填充的封闭边界
                Circle acCirc = new Circle();
                acCirc.Center = new Point3d(914837.7813, 534890.7717, 0);
                acCirc.Radius = 1000;
                // 将圆添加到块表记录和事务
                acBlkTblRec.AppendEntity(acCirc);
                acTrans.AddNewlyCreatedDBObject(acCirc, true);
                // 将圆的 ObjectId 添加到对象数组
                ObjectIdCollection acObjIdColl = new ObjectIdCollection();
                acObjIdColl.Add(acCirc.ObjectId);
                // 创建填充对象并添加到块表记录
                Hatch acHatch = new Hatch();
                acBlkTblRec.AppendEntity(acHatch);
                acTrans.AddNewlyCreatedDBObject(acHatch, true);
                // 设置填充对象的属性
                // 在调用 AppendLoop 之前设置填充对象的关联属性
                acHatch.SetHatchPattern(HatchPatternType.PreDefined, "HONEY");
                acHatch.Associative = true;
                acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl);
                acHatch.EvaluateHatch(true);
                // 将新对象保存到数据库
                acTrans.Commit();
            }
        }


        [CommandMethod("CheckForPickfirstSelection", CommandFlags.UsePickSet)]
        public static void CheckForPickfirstSelection()
        {
            // 获取当前文档获
            Editor acDocEd = Application.DocumentManager.MdiActiveDocument.Editor;
            // 获取 PickFirst 选择集
            PromptSelectionResult acSSPrompt;
            acSSPrompt = acDocEd.SelectImplied();
            SelectionSet acSSet;
            // 如果提示状态 OK，说明启动命令前选择了对象;
            if (acSSPrompt.Status == PromptStatus.OK)
            {
                acSSet = acSSPrompt.Value;
                Application.ShowAlertDialog("Number of objects in Pickfirst selection: " +
                acSSet.Count.ToString());
            }
            else
            {
                Application.ShowAlertDialog("Number of objects in Pickfirst selection: 0");
            }
            // 清空 PickFirst 选择集
            ObjectId[] idarrayEmpty = new ObjectId[0];
            acDocEd.SetImpliedSelection(idarrayEmpty);
            // 请求从图形区域选择对象
            acSSPrompt = acDocEd.GetSelection();
            // 如果提示状态 OK，表示已选择对象
            if (acSSPrompt.Status == PromptStatus.OK)
            {
                acSSet = acSSPrompt.Value;
                Application.ShowAlertDialog("Number of objects selected: " +
                acSSet.Count.ToString());
            }
            else
            {
                Application.ShowAlertDialog("Number of objects selected: 0");
            }
        }


        [CommandMethod("FilterBlueCircleOnLayer0")]
        public static void FilterBlueCircleOnLayer0()
        {
            //获取当前文档编辑器
            Editor acDocEd = Application.DocumentManager.MdiActiveDocument.Editor;
            // 创建 TypedValue 数组定义过滤条件
            TypedValue[] acTypValAr = new TypedValue[3];
            acTypValAr.SetValue(new TypedValue((int)DxfCode.Color, 5), 0);
            acTypValAr.SetValue(new TypedValue((int)DxfCode.Start, "HATCH"), 1);
            acTypValAr.SetValue(new TypedValue((int)DxfCode.LayerName, "0"), 2);
            // 将过滤条件赋值给 SelectionFilter 对象
            SelectionFilter acSelFtr = new SelectionFilter(acTypValAr);
            // 请求在图形区域选择对象
            PromptSelectionResult acSSPrompt;
            acSSPrompt = acDocEd.GetSelection(acSelFtr);
            // 如果提示状态 OK，表示对象已选
            if (acSSPrompt.Status == PromptStatus.OK)
            {
                SelectionSet acSSet = acSSPrompt.Value;
                Application.ShowAlertDialog("Number of objects selected: " +
                    acSSet.Count.ToString());
            }
            else
            {
                Application.ShowAlertDialog("Number of objects selected: 0");
            }
        }

        [CommandMethod("CreateRotatedDimension")]
        public static void CreateRotatedDimension()
        {
            // 获取当前数据库
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            // 启动事务
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // 以读模式打开块表
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                // 以写模式打开块表记录 ModelSpace（模型空间）
                BlockTableRecord acBlkTblRec;
                acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                OpenMode.ForWrite) as BlockTableRecord;
                // 创建旋转尺寸标注
                AlignedDimension acRotDim = new AlignedDimension();
                //acRotDim.XLine2Point
                acRotDim.XLine2Point = new Point3d(0, 0, 0);
                acRotDim.XLine2Point = new Point3d(6, 3, 0);
                //acRotDim.Rotation = 0.707;
                acRotDim.DimLinePoint = new Point3d(0, 5, 0);
                acRotDim.DimensionStyle = acCurDb.Dimstyle;
                // 将新对象添加到块表记录 ModelSpace 及事务
                acBlkTblRec.AppendEntity(acRotDim);
                acTrans.AddNewlyCreatedDBObject(acRotDim, true);
                // 提交修改，关闭事务
                acTrans.Commit();
            }
        }

    }

}
