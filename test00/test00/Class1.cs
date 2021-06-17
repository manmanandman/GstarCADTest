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
        public static String conn_string = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\\KMUTT\\GstarCADTest\\XdataTest.accdb;Persist Security Info= False";

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

        [CommandMethod("ViewXData")]
        public void ViewXData()
        {
            // Get the current database and start a transaction
            Database acCurDb;
            acCurDb = Application.DocumentManager.MdiActiveDocument.Database;

            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Editor edt = acDoc.Editor;

            // Prompt the user using PromptStringOptions
            PromptStringOptions prompt = new PromptStringOptions("Enter AppName : ");
            prompt.AllowSpaces = true;

            // Get the results of the user input using a PromptResult
            PromptResult promptResult = edt.GetString(prompt);

            //string appName = "AutoPost2004";
            string appName = promptResult.StringResult;

            string msgstr = "";

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
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
                            // Get the values in the xdata
                            foreach (TypedValue typeVal in rb)
                            {
                                msgstr = msgstr + "\n" + typeVal.TypeCode.ToString() + " : " + typeVal.Value;
                            }
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
            Editor edt = acDoc.Editor;

            // Prompt the user using PromptStringOptions
            PromptStringOptions prompt = new PromptStringOptions("Enter AppName : ");
            prompt.AllowSpaces = true;

            // Get the results of the user input using a PromptResult
            PromptResult promptResult = edt.GetString(prompt);

            //string appName = "AutoPost2004";
            string appName = promptResult.StringResult;
            //string appName = "AutoPost2004";

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

                            if (tv[1].TypeCode == 1070 && (tv[1].Value.ToString() == "102" || tv[1].Value.ToString() == "12"))
                            {
                                ClearDB();
                                CloseDatabase();
                                Application.ShowAlertDialog("Can't create report.\nThere is level 2 Tendon selected.");
                                return;
                            }
                            else if (tv[1].TypeCode == 1070 && (tv[1].Value.ToString() == "103" || tv[1].Value.ToString() == "13"))
                            {
                                tendon = 3;
                                tendon3num++;
                                tendonSubId = tendon3num;
                            }
                            else if (tv[1].TypeCode == 1070 && (tv[1].Value.ToString() == "104" || tv[1].Value.ToString() == "14"))
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

                    if(tendon3num!=0 && tendon4num!=0)
                    {
                        Form1 f = new Form1();
                        f.Show();
                        f.GetReport();
                    }
                    else
                    {
                        Application.ShowAlertDialog("No tendon level 3 or 4 found.");
                    }
                }

                // Ends the transaction and ensures any changes made are ignored
                acTrans.Abort();

                // Dispose of the transaction
            }
        }

        [CommandMethod("GetTendonResult")]
        public void GetTendonResult()
        {
            int tendon1num = 0;
            int tendon2num = 0;
            int tendon3num = 0;
            int tendon4num = 0;
            int elsenum = 0;

            // Get the current database and start a transaction
            Database acCurDb;
            acCurDb = Application.DocumentManager.MdiActiveDocument.Database;

            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Editor edt = acDoc.Editor;

            // Prompt the user using PromptStringOptions
            PromptStringOptions prompt = new PromptStringOptions("Enter AppName : ");
            prompt.AllowSpaces = true;

            // Get the results of the user input using a PromptResult
            PromptResult promptResult = edt.GetString(prompt);

            //string appName = "AutoPost2004";
            string appName = promptResult.StringResult;
            //string appName = "AutoPost2004";

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

                            if (tv[1].TypeCode == 1070 && (tv[1].Value.ToString() == "206" || tv[1].Value.ToString() == "11"))
                            {
                                tendon = 1;
                                tendon1num++;
                                tendonSubId = tendon1num;
                            }
                            else if (tv[1].TypeCode == 1070 && (tv[1].Value.ToString() == "102" || tv[1].Value.ToString() == "12"))
                            {
                                tendon = 2;
                                tendon2num++;
                                tendonSubId = tendon2num;
                            }
                            else if (tv[1].TypeCode == 1070 && (tv[1].Value.ToString() == "103" || tv[1].Value.ToString() == "13"))
                            {
                                tendon = 3;
                                tendon3num++;
                                tendonSubId = tendon3num;
                            }
                            else if (tv[1].TypeCode == 1070 && (tv[1].Value.ToString() == "104" || tv[1].Value.ToString() == "14"))
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

                    if (tendon3num != 0 && tendon4num != 0)
                    {
                        Form1 f = new Form1();
                        f.Show();
                        f.GetTotalTendon();
                    }
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
            Editor edt = acDoc.Editor;

            // Prompt the user using PromptStringOptions
            PromptStringOptions prompt = new PromptStringOptions("Enter AppName : ");
            prompt.AllowSpaces = true;

            // Get the results of the user input using a PromptResult
            PromptResult promptResult = edt.GetString(prompt);

            //string appName = "AutoPost2004";
            string appName = promptResult.StringResult;
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
                                msgstr = msgstr + "\n" + tv[order.Value-1].TypeCode.ToString() + " : " + tv[order.Value-1].Value;

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

        // Copy xData to new AppName
        [CommandMethod("CopyXdata")]
        public void CopyXdata()
        {
            // Get the current database and start a transaction
            Database acCurDb;
            acCurDb = Application.DocumentManager.MdiActiveDocument.Database;
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Editor edt = acDoc.Editor;

            // Prompt the user using PromptStringOptions
            PromptStringOptions prompt = new PromptStringOptions("Enter Old AppName : ");
            prompt.AllowSpaces = true;

            // Get the results of the user input using a PromptResult
            PromptResult promptResult = edt.GetString(prompt);

            // Prompt the user using PromptStringOptions
            PromptStringOptions prompt2 = new PromptStringOptions("Enter New AppName : ");
            prompt2.AllowSpaces = true;

            // Get the results of the user input using a PromptResult
            PromptResult promptResult2 = edt.GetString(prompt2);


            string appName = promptResult.StringResult;
            string newAppName = promptResult2.StringResult;

            string msgstr = "";

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
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
                                                         OpenMode.ForWrite) as Entity;

                        // Get the extended data attached to each object for MY_APP
                        ResultBuffer rb = acEnt.GetXDataForApplication(appName);

                        // Make sure the Xdata is not empty
                        if (rb != null)
                        {
                            try
                            {
                                // Change data into array
                                TypedValue[] tv = rb.AsArray();

                                // Change AppName to new name that store in first array 
                                tv[0] = new TypedValue((int)DxfCode.ExtendedDataRegAppName, newAppName);

                                // Chage value of tendon to 2022 style
                                if (tv[1].TypeCode == 1070)
                                {
                                    tv[1] = ChangeValue2022(tv[1]);
                                }




                                // Create result Buffer to add data to Xdata
                                using (ResultBuffer rb2 = new ResultBuffer())
                                {
                                    // Get the values in the xdata

                                    // Each TypedValue called "tvs" in array "tv"
                                    foreach (TypedValue tvs in tv)
                                    {
                                        rb2.Add(tvs);
                                    }
                                    acEnt.XData = rb2;
                                }

                            }
                            catch (System.Exception ex)
                            {
                                Application.ShowAlertDialog(ex.Message);
                                return;
                            }
                           

                        }

                        msgstr = "";
                    }
                }

                // Send the transaction
                acTrans.Commit();

                acDoc.Editor.WriteMessage("\nFinish copy xdata from " + appName + " to " + newAppName);
            }
        }

        TypedValue ChangeValue2022(TypedValue tv)
        {
            switch (tv.Value.ToString())
            {
                case "206":
                    return new TypedValue((int)DxfCode.ExtendedDataInteger16, "11");

                case "102":
                    return new TypedValue((int)DxfCode.ExtendedDataInteger16, "12");

                case "103":
                    return new TypedValue((int)DxfCode.ExtendedDataInteger16, "13");

                case "104":
                    return new TypedValue((int)DxfCode.ExtendedDataInteger16, "14")
                        ;
                case "201":
                    return new TypedValue((int)DxfCode.ExtendedDataInteger16, "21");

                case "202":
                    return new TypedValue((int)DxfCode.ExtendedDataInteger16, "22");

                case "203":
                    return new TypedValue((int)DxfCode.ExtendedDataInteger16, "23");

                case "204":
                    return new TypedValue((int)DxfCode.ExtendedDataInteger16, "24");

                case "205":
                    return new TypedValue((int)DxfCode.ExtendedDataInteger16, "25");

                default:
                    return tv;
            }
        }



        /*=================================================

                        DATABASE CODING SECTION

        =================================================== */


        // Declared connection variable for database section
        OleDbConnection conn = null;

        // Connect to Aceess DB (Database)
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

        // Disconnect
        public void CloseDatabase()
        {
            try
            {
                conn.Close();
            }
            catch (System.Exception ex) { Application.ShowAlertDialog(ex.Message); }
        }

        // Push data to DB
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

        // Clear data from DB
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
