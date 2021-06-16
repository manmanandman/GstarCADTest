using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using GrxCAD.Runtime;
using GrxCAD.ApplicationServices;
using GrxCAD.DatabaseServices;
using GrxCAD.EditorInput;
using GrxCAD.Geometry;
using GrxCAD.Colors;

using System.Data.OleDb;

[assembly: CommandClass(typeof(test00.HelloCmd))]

namespace test00
{
    public class HelloCmd 
    {
        [CommandMethod("DrawSquare")]
        public void DrawSquare()
        {
            // Get the document object
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // Create PromptIntegerOption
            PromptIntegerOptions pIntOpts = new PromptIntegerOptions("");

            // Restrict input to positive and non-negative values
            pIntOpts.AllowZero = false;
            pIntOpts.AllowNegative = false;

            // Get the value entered by the user
            pIntOpts.Message = "\nEnter the size of width";
            PromptIntegerResult width = doc.Editor.GetInteger(pIntOpts);

            pIntOpts.Message = "\nEnter the size of length";
            PromptIntegerResult lenght = doc.Editor.GetInteger(pIntOpts);

            // Promt for the starting point
            PromptPointOptions ppo = new PromptPointOptions("Pick starting point : ");
            PromptPointResult ppr = ed.GetPoint(ppo);

            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                try
                {
                    BlockTable bt;
                    bt = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;

                    // Open the block table record Model space for write
                    BlockTableRecord btr;
                    btr = trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    ed.WriteMessage("\nDrawing a Square:");
                    Point3d pt1 = new Point3d(ppr.Value.X, ppr.Value.Y, 0);
                    Point3d pt2 = new Point3d(ppr.Value.X + lenght.Value, ppr.Value.Y, 0);
                    Point3d pt3 = new Point3d(ppr.Value.X + lenght.Value, ppr.Value.Y + width.Value, 0);
                    Point3d pt4 = new Point3d(ppr.Value.X, ppr.Value.Y + width.Value, 0);

                    Line ln1 = new Line(pt1, pt2);
                    Line ln2 = new Line(pt2, pt3);
                    Line ln3 = new Line(pt3, pt4);
                    Line ln4 = new Line(pt4, pt1);

                    ln1.ColorIndex = 1;
                    ln2.ColorIndex = 2;
                    ln3.ColorIndex = 3;
                    ln4.ColorIndex = 4;

                    btr.AppendEntity(ln1);
                    btr.AppendEntity(ln2);
                    btr.AppendEntity(ln3);
                    btr.AppendEntity(ln4);

                    trans.AddNewlyCreatedDBObject(ln1, true);
                    trans.AddNewlyCreatedDBObject(ln2, true);
                    trans.AddNewlyCreatedDBObject(ln3, true);
                    trans.AddNewlyCreatedDBObject(ln4, true);

                    trans.Commit();
                }
                catch (System.Exception ex)
                {
                    doc.Editor.WriteMessage("Error encountered: " + ex.Message);
                    trans.Abort();
                }
            }
        }

        [CommandMethod("GetReport")]
        public void GetReport()
        {
            int tendon3num = 0;
            int tendon4num = 0;
            int elsenum = 0;

            // Get the current database and start a transaction
            Database acCurDb;
            acCurDb = Application.DocumentManager.MdiActiveDocument.Database;

            Document acDoc = Application.DocumentManager.MdiActiveDocument;

            string appName = "AutoPost2004";

            string msgstr = "";

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Request objects to be selected in the drawing area
                PromptSelectionResult acSSPrompt = acDoc.Editor.GetSelection();

                // If the prompt status is OK, objects were selected
                if (acSSPrompt.Status == PromptStatus.OK)
                {
                    SelectionSet acSSet = acSSPrompt.Value;


                    // Start connect to db
                    ConnectDatabase();
                    ClearDB();

                    // Step through the objects in the selection set
                    foreach (SelectedObject acSSObj in acSSet)
                    {
                        // Open the selected object for read
                        Entity acEnt = acTrans.GetObject(acSSObj.ObjectId,
                                                         OpenMode.ForRead) as Entity;

                        // Get the extended data attached to each object for MY_APP
                        ResultBuffer rb = acEnt.GetXDataForApplication(appName);

                        // Make sure the Xdata is not empty
                        if (rb != null)
                        {
                            TypedValue[] tv = rb.AsArray();
                            var tendon = 0;
                            var tendonSubId = 0;

                            if (tv[1].TypeCode == 1070 && tv[1].Value.ToString() == "102")
                            {
                                ClearDB();
                                CloseDatabase();
                                Application.ShowAlertDialog("Can't create report.\nThere is level 2 Tendon selected.");
                                return;
                            }
                            else if (tv[1].TypeCode == 1070 && tv[1].Value.ToString() == "103")
                            {
                                tendon = 3;
                                tendon3num++;
                                tendonSubId = tendon3num;
                            }
                            else if (tv[1].TypeCode == 1070 && tv[1].Value.ToString() == "104")
                            {
                                tendon = 4;
                                tendon4num++;
                                tendonSubId = tendon4num;
                            }
                            else
                            {
                                tendon = 0;
                                elsenum++;
                                tendonSubId = elsenum;
                            }

                            // Get the values in the xdata
                            foreach (TypedValue typeVal in rb)
                            {
                                msgstr = msgstr + "\n" + typeVal.TypeCode.ToString() + " : " + typeVal.Value;
                                AddDataToDatabase(tendon, tendonSubId, typeVal.TypeCode.ToString(), typeVal.Value.ToString());
                            }
                        }
                        else
                        {
                            msgstr = "NONE";
                        }

                        // Display the values returned
                        //Application.ShowAlertDialog(appName + " xdata on " + acEnt.GetType().ToString() + ":\n" + msgstr);


                        msgstr = "";
                    }
                    CloseDatabase();
                    form();
                }

                // Ends the transaction and ensures any changes made are ignored
                acTrans.Abort();

                // Dispose of the transaction
            }
        }

        [CommandMethod("getEdge")]
        public static void Edge()
        {

            Document Doc = Application.DocumentManager.MdiActiveDocument;
            Database Db = Doc.Database;
            Editor edt = Doc.Editor;


            string appName = "AutoPost2004";
            string xdataStr = "201";


            using (Transaction Trans = Db.TransactionManager.StartTransaction())
            {
                PromptSelectionResult psr = edt.GetSelection();
                if (psr.Status == PromptStatus.OK)
                {
                    RegAppTable rat = (RegAppTable)Trans.GetObject(Db.RegAppTableId, OpenMode.ForRead, false); ;

                    if (rat.Has(appName) == false)
                    {
                        using (RegAppTableRecord ratr = new RegAppTableRecord())
                        {
                            ratr.Name = appName;

                            rat.UpgradeOpen();
                            rat.Add(ratr);
                            Trans.AddNewlyCreatedDBObject(ratr, true);
                        }
                    }

                    SelectionSet ss = psr.Value; //Create Selectionset and store object in selectionset ss

                    //open the linetype table for read                            
                    LinetypeTable ltr;
                    ltr = Trans.GetObject(Db.LinetypeTableId, OpenMode.ForRead) as LinetypeTable;

                    //open the blocktable for read
                    BlockTable bt;
                    bt = Trans.GetObject(Db.BlockTableId, OpenMode.ForRead) as BlockTable;

                    //open the block table recoed model spae for write
                    BlockTableRecord btr;
                    btr = Trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    string sLineTypName = "Center";

                    //check line type on Autocad
                    if (ltr.Has(sLineTypName) == false)
                    {
                        Db.LoadLineTypeFile(sLineTypName, "acad.lin");
                    }

                    // Prompt the user using PromptStringOptions
                    PromptIntegerOptions thick = new PromptIntegerOptions("Enter thickness ");
                    thick.AllowNegative = false;
                    thick.AllowZero = false;
                    PromptIntegerResult thickness = edt.GetInteger(thick);

                    int x = thickness.Value;


                    int index = 65;

                    //check vertical
                    if (true) //แนวตั้ง
                    {
                        string name = (Convert.ToChar(index)).ToString(); //convert ascii to string
                        using (ResultBuffer rb = new ResultBuffer())
                        {
                            {
                                rb.Add(new TypedValue((int)DxfCode.ExtendedDataRegAppName, appName)); //appname 1001
                                rb.Add(new TypedValue((int)DxfCode.ExtendedDataInteger16, xdataStr)); //Gridline 1070
                                rb.Add(new TypedValue((int)DxfCode.ExtendedDataReal, x)); // name 1040


                                // Step through the objects in the selection set
                                foreach (SelectedObject sObj in ss)
                                {
                                    // Open the selected object for write
                                    Entity Ent = Trans.GetObject(sObj.ObjectId, OpenMode.ForWrite) as Entity;


                                    // Append the extended data to each object
                                    Ent.XData = rb;

                                    Doc.Editor.WriteMessage("\n" + Ent.XData);


                                    //change line type and color index
                                    Ent.Linetype = sLineTypName;
                                    Ent.ColorIndex = 1;
                                }
                            }
                        }
                    }

                    LayerTable lyTab = Trans.GetObject(Db.LayerTableId, OpenMode.ForRead) as LayerTable;
                    if (lyTab.Has("Edge"))
                    {
                        Doc.Editor.WriteMessage("Layer already exist.\n");
                        //Trans.Abort();
                    }
                    else
                    {
                        lyTab.UpgradeOpen();
                        LayerTableRecord lt = new LayerTableRecord();
                        lt.Name = "Edge";
                        lt.Color = Color.FromColorIndex(ColorMethod.ByLayer, 1);
                        lyTab.Add(lt);
                        Trans.AddNewlyCreatedDBObject(lt, true);
                        Db.Clayer = lyTab["Edge"];

                        Doc.Editor.WriteMessage("Layer [" + lt.Name + "] was created successfully.\n");

                        // Commit the transaction
                    }

                    // Move Object to Edge layer
                    if (lyTab.Has("Edge"))
                    {
                        int changedCount = 0;

                        // We have the layer table open, so let's get the layer ID and use that
                        ObjectId lid = lyTab["Edge"];

                        foreach (ObjectId id in psr.Value.GetObjectIds())
                        {

                            Entity ent = (Entity)Trans.GetObject(id, OpenMode.ForWrite);

                            ent.LayerId = lid;

                            // Could also have used:

                            //  ent.Layer = newLayerName;

                            // but this way is more efficient and cleaner

                            changedCount++;

                            Doc.Editor.WriteMessage("total change : " + changedCount.ToString()+"\n");
                        }

                    }
                }

                Trans.Commit();


            }
        }

        [CommandMethod("PeekXdata")]
        public void PeekXdata()
        {
            // Get the current database and start a transaction
            Database acCurDb;
            acCurDb = Application.DocumentManager.MdiActiveDocument.Database;

            Document acDoc = Application.DocumentManager.MdiActiveDocument;

            string appName = "AutoPost2004";

            string msgstr = "";

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Create PromptIntegerOption
                PromptIntegerOptions pIntOpts = new PromptIntegerOptions("");

                // Restrict input to positive and non-negative values
                pIntOpts.AllowZero = false;
                pIntOpts.AllowNegative = false;

                // Get the value entered by the user
                pIntOpts.Message = "\nEnter the order to peek";
                PromptIntegerResult order = acDoc.Editor.GetInteger(pIntOpts);

                // Request objects to be selected in the drawing area
                PromptSelectionResult acSSPrompt = acDoc.Editor.GetSelection();

                // If the prompt status is OK, objects were selected
                if (acSSPrompt.Status == PromptStatus.OK)
                {
                    SelectionSet acSSet = acSSPrompt.Value;

                    // Step through the objects in the selection set
                    foreach (SelectedObject acSSObj in acSSet)
                    {
                        // Open the selected object for read
                        Entity acEnt = acTrans.GetObject(acSSObj.ObjectId,
                                                         OpenMode.ForRead) as Entity;

                        // Get the extended data attached to each object for MY_APP
                        ResultBuffer rb = acEnt.GetXDataForApplication(appName);

                        // Make sure the Xdata is not empty
                        if (rb != null)
                        {
                            TypedValue[] tv = rb.AsArray();
                            try
                            {
                                msgstr = msgstr + "\n" + tv[order.Value].TypeCode.ToString() + " : " + tv[order.Value].Value;

                            }
                            catch
                            {
                                
                            }

                            //// Get the values in the xdata
                            //foreach (TypedValue typeVal in rb)
                            //{
                            //    msgstr = msgstr + "\n" + typeVal.TypeCode.ToString() + " : " + typeVal.Value;
                            //}
                        }
                        else
                        {
                            msgstr = "NONE";
                        }

                        // Display the values returned
                        Application.ShowAlertDialog(appName + " xdata on " + acEnt.GetType().ToString() + ":\n" + msgstr);


                        msgstr = "";
                    }
                }

                // Ends the transaction and ensures any changes made are ignored
                acTrans.Abort();

                // Dispose of the transaction
            }
        }

        [CommandMethod("openform")]
        public void form()
        {
            Form1 f = new Form1();
            f.Show();
        }

        public static String conn_string = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\\KMUTT\\GstarCADTest\\XdataTest.accdb;Persist Security Info= False";
        OleDbConnection conn = null;

        public void ConnectDatabase()
        {
            try
            {
                conn = new OleDbConnection(conn_string);
                conn.Open();
            }
            catch (System.Exception ex)
            {
                Application.ShowAlertDialog(ex.Message);
            }
        }

        public void CloseDatabase()
        {
            try
            {
                conn.Close();
            }
            catch (System.Exception ex) { Application.ShowAlertDialog(ex.Message); }
        }

        public void AddDataToDatabase(int tendon,int tendonSubId,string typecode, string values)
        {
            string q = "INSERT INTO XDATA_TEST VALUES ('" + tendon + "','" + tendonSubId + "','" + typecode + "','" + values + "')";
            if(tendon != 0)
            {
                try
                {
                    OleDbCommand command = new OleDbCommand();
                    command.Connection = conn;
                    command.CommandText = q;
                    command.ExecuteNonQuery();
                }
                catch (System.Exception ex) { Application.ShowAlertDialog(ex.Message); }
            }
        }

        public void ClearDB()
        {
            string q = "DELETE * FROM XDATA_TEST";
            try
            {
                OleDbCommand command = new OleDbCommand();

                command.Connection = conn;
                command.CommandText = q;
                command.ExecuteNonQuery();
            }
            catch (System.Exception ex) { Application.ShowAlertDialog(ex.Message); }
        }
    }


}
