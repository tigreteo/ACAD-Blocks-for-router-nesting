using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;

using System;
using System.Text;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using System.Data;
using System.Globalization;

//TODO:
//CHECK THAT THE PART NUMBER LISTED LINES UP WITH SAVE NAME
//CREATE A WINDOW FOR ERROR LOG OR ALL CLEAR MESSAGE AFTER CREATING PARTS 
//-format the stringbuilder errors

namespace Blocks_for_Nesting
{
    public class Main
    {
        //global used to hold a log of drawing errors, not programmatic errors
        public static StringBuilder partsErrorLog = new StringBuilder();

        [CommandMethod("PrepToNest")]
        public void RedefiningABlock()
        {
            partsErrorLog.Clear();
            // Get the current database and start a transaction
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                //zoom out to make sure the BTR can see the whole project
                Zoom.ZoomExt(new Point3d(), new Point3d(), new Point3d(), 1);
        
                // Open the Block table for read
                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                //method to search drawing for a block that would hold the parts and delete them all                              
                cleanSlate(tr, bt, btr, ed);                

                //recursive purge the old blkref data
                using (Transaction purgeTr = db.TransactionManager.StartTransaction())
                {
                    bool contPurge = true;
                    while (contPurge == true)
                    { contPurge = purgeAll(db, purgeTr, ed, bt); }

                    purgeTr.Commit();
                }                
                //add a save to help keep BTRs to work for modififying the drawing
                db.SaveAs(doc.Name, true, DwgVersion.Current, db.SecurityParameters);
                             
                ObjectIdCollection affectedParts = new ObjectIdCollection();
                #region select parts
                //selects all parts now, was filtering just some parts for testing 
                //old technique of selecting everything on screen
                //was grabbing blocks that were deleted and purged
                //foreach (ObjectId msId in btr)
                //{
                //    if(msId.ObjectClass.DxfName.ToUpper() == "INSERT")
                //    { affectedParts.Add(msId); }
                //}

                // Create a TypedValue array to define the filter criteria
                TypedValue[] filterList = new TypedValue[1];
                filterList.SetValue(new TypedValue(0, "INSERT"), 0);

                SelectionFilter setFilter = new SelectionFilter(filterList);

                //get selection of area to be affected
                SelectionSet acSSet;
                PromptSelectionResult selection = ed.GetSelection(setFilter);
                //only proceed if the selection has been made
                if (selection.Status == PromptStatus.OK)
                { acSSet = selection.Value; }
                else
                { return; }

                affectedParts = new ObjectIdCollection(acSSet.GetObjectIds());
                #endregion

                //zoom to affected parts
                Extents3d zoomTarget = TransactionExtensions.GetExtents(tr, affectedParts);
                Zoom.ZoomExt(zoomTarget.MinPoint, zoomTarget.MaxPoint, new Point3d(), 1.1);
                

                //get a point to use later when we wrap the parts in a border
                var extMax = db.Extmax;
                double varX = extMax.X + 10;
                double varY = extMax.Y - 10;
                double YoffSet = 0;

                //find expected styleID and group number from filename
                #region find name
                string fileName = Path.GetFileNameWithoutExtension(doc.Name);
                StringBuilder group = new StringBuilder();
                StringBuilder style = new StringBuilder();
                bool secondPart = false;
                bool firstPart = false;
                char[] groupList = { ' ', '-','A','C' };
                char[] styleList = { ' ', 'S','N'};//might need to change this to be more general
                //it assumes and S as in SAM1, or N as in NESTERWOOD

                //loop through file name to generate parts of name
                //first to last
                foreach (char c in fileName)
                {
                    //if number is false it isnt complete
                    if (!secondPart)
                    {
                        if (!firstPart)
                        {
                            //if part isnt one of the separaters add it to end of part
                            if (Array.Exists(groupList, element => element == c))
                            {
                                //if it is a separater then the first part is complete(true)
                                firstPart = true;
                            }
                            else
                                group.Append(c);
                        }

                        else {
                            //if par isnt one of the expected closers then add it to the end of the part
                            if (Array.Exists(styleList, element => element == c))
                            {
                                //if it is an expected closer then set second part to true
                                secondPart = true;
                            }
                            else
                                style.Append(c);
                        }
                    }
                }
                #endregion

                //loop through list to remove any blocks that arent a border
                foreach (ObjectId objId in affectedParts)
                {
                    BlockReference blkRef = tr.GetObject(objId, OpenMode.ForRead) as BlockReference;
                    string blockName = "";
                    //if there are any undefined blocks around, then they wont have names to look for
                    try
                    { blockName = blkRef.Name; }
                    catch (System.Exception)
                    { }
                    if (!blockName.Contains("CNC_BORDER"))
                        affectedParts.Remove(objId);
                }

                //check that all the cnc borders don't overlap
                #region check overlap
                bool overlapped = false;
                foreach(ObjectId msId in affectedParts)
                {
                    foreach(ObjectId secondId in affectedParts)
                    {
                        if (msId != secondId)
                        {
                            bool results = plIntersect(msId, secondId, db);
                            if (results)
                                overlapped = true;
                        }
                    }
                }
                if (overlapped)
                {
                    //identify offending overlapp
                    ed.WriteMessage("\nERROR: Overlapping borders");
                    partsErrorLog.Append("Critical: Overlapping borders");
                    //locate where the error is
                    //dispaly error list, bc this error breaks the command
                    warnBox();
                    return;
                }
                #endregion

                bool validBlocks = true;
                //loop through the list to create blks in a designated area
                foreach(ObjectId objId in affectedParts)
                {
                    BlockReference blkRef = tr.GetObject(objId, OpenMode.ForRead) as BlockReference;
                    string blockName = "";

                    //if there are any undefined blocks around, then they wont have names to look for
                    try                    
                       { blockName = blkRef.Name; }
                    catch (System.Exception)
                    { }
                    
                    if(blockName.Contains("CNC_BORDER"))
                    {
                        //get a name for the blockref
                        //read attributes to generate a name
                        //should an attribute be missing, either request from user or generate and answer
                        string blkName = getName(blkRef, tr, group, style);

                        //find the quanitity of the part
                        double quantity = getQuantity(blkRef, tr);

                        //get coords of the blkref
                        Extents3d window = (Extents3d)blkRef.Bounds;

                        //get a selection set of the window
                        SelectionSet select;
                        PromptSelectionResult res = ed.SelectCrossingWindow(window.MinPoint, window.MaxPoint);
                        if(res.Status == PromptStatus.OK)
                            select = res.Value;
                        else
                            return;

                        ObjectIdCollection objIdCol = new ObjectIdCollection(select.GetObjectIds());
                        //remove CNC frame from that collection
                        objIdCol.Remove(objId);

                        //explode any blocks in the collection to be sure to be just looking at appropriate parts
                        objIdCol = explodeBlock(tr, objIdCol, btr);

                        ObjectIdCollection partCol = new ObjectIdCollection();
                        //need to sort out all of the dimension lines
                        //only keeping polylines & drill cirlces
                        foreach(ObjectId msId in objIdCol)
                        {
                            if(msId.ObjectClass.DxfName.ToUpper() == "POLYLINE" ||
                                msId.ObjectClass.DxfName.ToUpper() == "LWPOLYLINE" ||
                                msId.ObjectClass.DxfName.ToUpper() == "CIRCLE")
                            { partCol.Add(msId); }
                        }

                        //verify that something was selected
                        if(partCol.Count <= 0)
                            {
                            ed.WriteMessage("\n No parts found in part " + blkName +"\n");
                            partsErrorLog.AppendLine("Critical: No parts found in part " + blkName + "\n");
                            //dispaly error list, bc this error breaks the command
                            warnBox();
                            return;
                            }

                        //Run through checklist to make sure parts wont throw errors for NesterWood
                        //return a bool but not really using it. may use this to prevent the blocks from being formed
                        bool cont = verifyParts(partCol,btr,tr,ed,db, blkName);
                        if (!cont)
                            validBlocks = false; 

                        //clone the object Id Collection so that the orignal doesnt get blocked
                        ObjectIdCollection partCol2 = cloneCollection(partCol, tr, btr);

                        //loop through the parts to find the dist of X & Y for them
                        Extents3d extParts = getExtents(tr, partCol2);
                        double xdist = Math.Abs(extParts.MaxPoint.X - extParts.MinPoint.X);
                        double ydist = Math.Abs(extParts.MaxPoint.Y - extParts.MinPoint.Y);

                        //verify if the fileName is Valid
                        #region verify fileName is legal/valid
                        if (blkName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
                        {
                            //give warning about invalid name
                            ed.WriteMessage("\nFilename " + blkName + " uses invalid characters.\n");
                            partsErrorLog.AppendLine("Critical: Filename " + blkName + " uses invalid characters.\n");
                            //dispaly error list, bc this error breaks the command
                            warnBox();
                            return;
                        }
                        #endregion

                        //turn collection into a block reference
                        reblock(blkName, bt, tr, db, partCol, window, ydist, varX, varY, quantity, extParts);

                        //use the distX to offset part from last part plus some buffer space
                        varX = varX + xdist + 10;

                        //track and compare the min y position
                        if ((quantity + 1) * ydist > YoffSet)
                        { YoffSet = ((quantity + 1) * ydist) + 10; }
                    }
                }  
                //create a block with a reliable name
                //block is a rectangle from (oldMaxX, oldMaxY) to (newMaxX, NewMinY)
                surroundParts(tr, bt, btr, db, extMax, varX, YoffSet + 10);

                // Save the new object to the database
                tr.Commit();
                // Dispose of the transaction
            }

            //need to remove or at least flag all the old blocks that might be unreferenced now that they have been deleted
            //although these blocks delete themselves after opening again or if they are stacked
            //clearUnrefedBlocks(doc, ed, db); //seems to be deleting everything EXCEPT the undefined blocks

            //show all errors that might have not broken the command but could be issues
            warnBox();
        }

        [CommandMethod("PrepToEdit")]
        public void ChangeEditStyle()
        {
            // Get the current database and start a transaction
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            //create datatable to collate info of blocks
            //Block name, partname, part quantity, spec block id used, scale of spec
            System.Data.DataTable table = createTableHeader();

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                //zoom out to make sure the BTR can see the whole project
                Zoom.ZoomExt(new Point3d(), new Point3d(), new Point3d(), 1);

                // Open the Block table for read
                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                //TODO - move to standalone method for re-use
                ObjectIdCollection affectedParts = new ObjectIdCollection();
                #region select parts
                // Create a TypedValue array to define the filter criteria
                TypedValue[] filterList = new TypedValue[1];
                filterList.SetValue(new TypedValue(0, "INSERT"), 0);

                SelectionFilter setFilter = new SelectionFilter(filterList);

                //get selection of area to be affected
                SelectionSet acSSet;
                PromptSelectionResult selection = ed.GetSelection(setFilter);
                //only proceed if the selection has been made
                if (selection.Status == PromptStatus.OK)
                { acSSet = selection.Value; }
                else
                { return; }

                affectedParts = new ObjectIdCollection(acSSet.GetObjectIds());
                #endregion                

                ObjectIdCollection duplicateParts = new ObjectIdCollection();
                //load blocks into datatable with part name and quantity
                foreach (ObjectId objId in affectedParts)
                {
                    BlockReference blkRef = tr.GetObject(objId, OpenMode.ForRead) as BlockReference;

                    //find the name from the blkref name
                    string blockName = "";
                    try
                    { blockName = blkRef.Name; }
                    catch (System.Exception)
                    { break; }

                    //try to add to table
                    //if already exists, add to quantity, remove from collection
                    DataRow[] foundEntry = table.Select("BlockName = '" + blockName + "'");
                    if (foundEntry.Length == 0)
                    {
                        //add styleID to table
                        DataRow newRow = table.NewRow();
                        newRow["BlockName"] = blockName;
                        newRow["Quantity"] = 1;
                        newRow["BlockIdPart"] = blkRef.Handle;

                        //following expected convention should yield name for when we insert block ref
                        //split name by - to get parts?
                        //often name is separated by part number by blank space
                        string[] nameParts = blockName.Split('-');
                        string partName = nameParts[nameParts.Length - 1];

                        newRow["PartName"] = partName;

                        //add new row to table
                        table.Rows.Add(newRow);
                    }
                    else
                    {
                        //load the first data row (should only ever be qty of 1)
                        DataRow updateRow = foundEntry[0];
                        double qty = Convert.ToDouble(updateRow["Quantity"]);
                        //update quantity
                        updateRow["Quantity"] = qty + 1;

                        //remove id from collection
                        duplicateParts.Add(objId);
                        //affectedParts.Remove(objId);
                    }

                }

                foreach (ObjectId objId in duplicateParts)
                {
                    //remove parts without throwing off sequence
                    affectedParts.Remove(objId);
                }

                //insert spec around each block
                //input info from table
                //add scale and blockid to table
                foreach (ObjectId objId in affectedParts)
                {
                    BlockReference blkRef = tr.GetObject(objId, OpenMode.ForRead) as BlockReference;
                    //find the name from the blkref name
                    string blockName = "";
                    try
                    { blockName = blkRef.Name; }
                    catch (System.Exception)
                    { break; }

                    //try to add to table
                    //if already exists, add to quantity, remove from collection
                    DataRow[] foundEntry = table.Select("BlockName = '" + blockName + "'");
                    if (foundEntry.Length > 0)
                    { Insert.insertSpec(blkRef, foundEntry[0]); }
                }

                tr.Commit();
            }


            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                //request start point for aligning all spec blocks
                Point3d startPartPoint = new Point3d();
                PromptPointResult ppr;
                PromptPointOptions ppo = new PromptPointOptions("Choose point to arrange parts from");
                ppr = ed.GetPoint(ppo);

                //if no point was given, just end at this point and user can still manually place parts
                if (ppr.Status == PromptStatus.OK)
                    startPartPoint = ppr.Value;
                else
                    return;

                //sort each around based on scale
                //arrange each block w/ spec in order of largest to smallest
                DataView dv = table.DefaultView;
                dv.Sort = "Scale desc";
                table = dv.ToTable();

                startPartPoint = new Point3d(startPartPoint.X, startPartPoint.Y - 2, startPartPoint.Z);
                double scale = 0;
                Point3d lastPoint = new Point3d();
                //place each spec/part into location
                //each scale gets a separate row
                //each spec is spaced apart by 2
                foreach (DataRow r in table.Rows)
                {
                    double specScale = Convert.ToDouble(r["Scale"]);
                    if (scale == specScale)
                    {
                        lastPoint = moveSpec(db, r, lastPoint);
                    }
                    else
                    {
                        lastPoint = moveSpec(db, r, startPartPoint, true);
                        startPartPoint = new Point3d(startPartPoint.X, lastPoint.Y, 0);
                        scale = specScale;
                    }

                }
                tr.Commit();
            }
        }

        [CommandMethod("MirrorAllParts")]
        public void SwitchHands()
        {
            // Get the current database and start a transaction
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            //TODO - move to standalone method for re-use
            ObjectIdCollection affectedParts = new ObjectIdCollection();
            #region select parts
            // Create a TypedValue array to define the filter criteria
            TypedValue[] filterList = new TypedValue[1];
            filterList.SetValue(new TypedValue(0, "INSERT"), 0);

            SelectionFilter setFilter = new SelectionFilter(filterList);

            //get selection of area to be affected
            SelectionSet acSSet;
            PromptSelectionResult selection = ed.GetSelection(setFilter);
            //only proceed if the selection has been made
            if (selection.Status == PromptStatus.OK)
            { acSSet = selection.Value; }
            else
            { return; }

            affectedParts = new ObjectIdCollection(acSSet.GetObjectIds());
            #endregion

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                //loop through list to remove any blocks that arent a border
                #region remove to just Frame specs
                foreach (ObjectId objId in affectedParts)
                {
                    BlockReference blkRef = tr.GetObject(objId, OpenMode.ForRead) as BlockReference;
                    string blockName = "";
                    //if there are any undefined blocks around, then they wont have names to look for
                    try
                    { blockName = blkRef.Name; }
                    catch (System.Exception)
                    { }
                    if (!blockName.Contains("CNC_BORDER"))
                        affectedParts.Remove(objId);
                }
                #endregion

                //check that all the cnc borders don't overlap
                #region check overlap
                bool overlapped = false;
                foreach (ObjectId msId in affectedParts)
                {
                    foreach (ObjectId secondId in affectedParts)
                    {
                        if (msId != secondId)
                        {
                            bool results = plIntersect(msId, secondId, db);
                            if (results)
                                overlapped = true;
                        }
                    }
                }
                if (overlapped)
                {
                    //identify offending overlapp
                    ed.WriteMessage("\nERROR: Overlapping borders");
                    partsErrorLog.Append("Critical: Overlapping borders");
                    //locate where the error is
                    //dispaly error list, bc this error breaks the command
                    warnBox();
                    return;
                }
                #endregion

                //for each part                
                foreach (ObjectId objId in affectedParts)
                {
                    BlockReference blkRef = tr.GetObject(objId, OpenMode.ForRead) as BlockReference;
                    string blockName = "";

                    //if there are any undefined blocks around, then they wont have names to look for
                    try
                    { blockName = blkRef.Name; }
                    catch (System.Exception)
                    { break; }

                    #region rename any LH|RH parts
                    //re-name any LH <-> RH in partName
                    AttributeCollection attCol = blkRef.AttributeCollection;
                    bool updateName = false;
                    string[] nameParts = blockName.Split(' ');
                    for(int i = 0;i<nameParts.Length;i++)
                    {
                        if (nameParts[i] == "LH" || nameParts[i].ToUpper() == "LEFT")
                        {
                            nameParts[i] = "RH";
                            updateName = true;
                        }
                        else if (nameParts[i] == "RH" || nameParts[i].ToUpper() == "RIGHT")
                        {
                            nameParts[i] = "LH";
                            updateName = true;
                        }
                    }
                    if (updateName == true)
                    {
                        string partName = "";
                        foreach(string s in nameParts)
                        {
                            partName = partName + s + " ";
                        }
                        partName.Trim();
                        foreach (ObjectId objID in attCol)
                        {
                            AttributeReference acAttRef = tr.GetObject(objID, OpenMode.ForWrite) as AttributeReference;
                            if (acAttRef.Tag == "PARTNAME")
                            {
                                acAttRef.TextString = partName;
                            }
                        }
                    }
                    #endregion

                    if (blockName.Contains("CNC_BORDER"))
                    {
                        //get selection of everything inside of spec//get coords of the blkref
                        #region selection of specs inside spec
                        Extents3d window = (Extents3d)blkRef.Bounds;

                        //get a selection set of the window
                        SelectionSet select;
                        PromptSelectionResult res = ed.SelectCrossingWindow(window.MinPoint, window.MaxPoint);
                        if (res.Status == PromptStatus.OK)
                            select = res.Value;
                        else
                            return;

                        ObjectIdCollection objIdCol = new ObjectIdCollection(select.GetObjectIds());
                        //remove CNC frame from that collection
                        objIdCol.Remove(objId);
                        #endregion

                        //mid x of mirror
                        double midX = ((window.MaxPoint.X - window.MinPoint.X) / 2) + window.MinPoint.X;

                        //**ensure the system variable for mirroring text
                        //MIRRTEXT = 0

                        //!!**Mirror is putting curves into bottom view perspective (normal z = -1)
                        //

                        //mirror everything along center of spec, delete old content                        
                        foreach(ObjectId id in objIdCol)
                        {
                            Entity ent = tr.GetObject(id, OpenMode.ForWrite) as Entity;
                            Matrix3d mat = new Matrix3d();
                            ent.TransformBy(Matrix3d.Mirroring(new Line3d(
                                new Point3d(midX, 0, 0),
                                new Point3d(midX, 2, 0))));
                            //ent.TransformBy(Matrix3d.PlaneToWorld(new Plane()));
                            ent.TransformBy(Matrix3d.WorldToPlane(new Vector3d(0, 0, 1)));
                        }
                    }
                }
                tr.Commit();
            }
        }

        //checks to see if any of the blkrefs are overlapping, to make sure that the blocking methods wont grab other parts
        //res == null : ERROR
        //res < 0 no intersecting 
        //res > 0 intersecting 
        private static bool plIntersect(ObjectId obj0, ObjectId obj1, Database db)
        {
            double? area = null;
            BlockReference blk0 = null;
            BlockReference blk1 = null;
            Region re0 = new Region();

            using (Transaction tr = db.TransactionManager.StartTransaction())
            using (DBObjectCollection csc0 = new DBObjectCollection())
            using (DBObjectCollection csc1 = new DBObjectCollection())
            {
                blk0 = tr.GetObject(obj0, OpenMode.ForRead) as BlockReference;
                blk1 = tr.GetObject(obj1, OpenMode.ForRead) as BlockReference;
               
                Polyline pl0 = blkToPolyline(blk0);
                Polyline pl1 = blkToPolyline(blk1);
                try
                {
                    pl0.Explode(csc0);
                    pl1.Explode(csc1);
                    using (DBObjectCollection rec0 = Region.CreateFromCurves(csc0))
                    using (DBObjectCollection rec1 = Region.CreateFromCurves(csc1))
                    {
                        if (rec0.Count == 1 && rec1.Count == 1)
                        {
                            re0 = (Region)rec0[0];
                            re0.BooleanOperation(BooleanOperationType.BoolIntersect, (Region)rec1[0]);
                            area = new double?(re0.Area);
                        }
                    }
                }
                catch
                {
                }
                tr.Commit();
            }
            if(area > 0)
            {
                //find center point of intersection
                Extents3d bounds = re0.GeometricExtents;
                Point2d center = new Point2d(bounds.MinPoint.X + (bounds.MaxPoint.X - bounds.MinPoint.X) /2, bounds.MinPoint.Y + (bounds.MaxPoint.Y - bounds.MinPoint.Y)/2);
                //identify offending blockrefs
                //locate intersection by surrounding it with visible indicator
                partsErrorLog.AppendLine(blk0.Name + " intersecting" + blk1.Name + " at " + center);
                //ed.WriteMessage("/n"+blk0.Name + " intersecting" + blk1.Name + " at " + center); 

                //doesn't work bc nested transation rolls back any insertion
                #region surround intersection                
                //using (Transaction tr = db.TransactionManager.StartOpenCloseTransaction())
                //{
                //    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                //    BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                //    using (Polyline rectang = new Polyline())
                //    {
                //        rectang.AddVertexAt(0, new Point2d(bounds.MinPoint.X, bounds.MinPoint.Y), 0, 0, 0);
                //        rectang.AddVertexAt(1, new Point2d(bounds.MinPoint.X, bounds.MaxPoint.Y), 0, 0, 0);
                //        rectang.AddVertexAt(2, new Point2d(bounds.MaxPoint.X, bounds.MaxPoint.Y), 0, 0, 0);
                //        rectang.AddVertexAt(3, new Point2d(bounds.MaxPoint.X, bounds.MinPoint.Y), 0, 0, 0);
                //        rectang.Closed = true;
                //        rectang.Color = Color.FromRgb(23, 54, 232);

                //        btr.AppendEntity(rectang);
                //        tr.AddNewlyCreatedDBObject(rectang, true);
                //    }
                //    tr.Commit();
                //}
                #endregion

                return true;
            }
            else
            return false;
        }

        //create a polyline based on the extents of the blockref
        private static Polyline blkToPolyline(BlockReference blkRef)
        {
            Extents3d ext = (Extents3d)blkRef.Bounds;
            Polyline pLine = new Polyline();
            pLine.AddVertexAt(0, new Point2d(ext.MinPoint.X, ext.MinPoint.Y), 0, 0, 0);
            pLine.AddVertexAt(1, new Point2d(ext.MinPoint.X, ext.MaxPoint.Y), 0, 0, 0);
            pLine.AddVertexAt(2, new Point2d(ext.MaxPoint.X, ext.MaxPoint.Y), 0, 0, 0);
            pLine.AddVertexAt(3, new Point2d(ext.MaxPoint.X, ext.MinPoint.Y), 0, 0, 0);
            pLine.Closed = true;

            return pLine;
        }

        //loop through the collection of IDs to verify the parts collection before copying it to a block
        //Checking:
        // block will be less than 95 1/4" long
        // no duplicate polylines
        // no open poly lines
        private static bool verifyParts(ObjectIdCollection partCol,BlockTableRecord btr,Transaction tr, Editor ed, Database db, string blkName)
        {
            bool valid = true;
            //check length of part
            Extents3d selectionExt = tr.GetExtents(partCol);
            double length = Math.Abs(selectionExt.MaxPoint.X - selectionExt.MinPoint.X);
            length = length / 200;
            int scale = (int)length;
            if (scale == 0)
            { scale = 1; }
            if (Math.Abs(selectionExt.MaxPoint.X - selectionExt.MinPoint.X) > 95.25)
            {
                //create a dimensions of the object and write to console
                using (AlignedDimension dimlin = new AlignedDimension())
                {
                    dimlin.XLine1Point = new Point3d(selectionExt.MinPoint.X, selectionExt.MaxPoint.Y, 0);
                    dimlin.XLine2Point = new Point3d(selectionExt.MaxPoint.X, selectionExt.MaxPoint.Y, 0);
                    dimlin.DimensionStyle = db.Dimstyle;
                    dimlin.DimLinePoint = new Point3d(0, selectionExt.MaxPoint.Y + (8 * scale), 0);
                    dimlin.Dimtxt = 1 + scale;

                    btr.AppendEntity(dimlin);
                    tr.AddNewlyCreatedDBObject(dimlin, true);
                    ed.WriteMessage("\nError: Piece(s) too long in part " + blkName);
                    partsErrorLog.AppendLine("Warning: Piece(s) too long in part " + blkName);
                }
                    valid = false;
            }

            //look for open polylines
            foreach(ObjectId msId in partCol)
            {
                if (//msId.ObjectClass.DxfName.ToUpper() == "POLYLINE" ||
                    msId.ObjectClass.DxfName.ToUpper() == "LWPOLYLINE")
                {
                    Polyline pLine = tr.GetObject(msId, OpenMode.ForRead) as Polyline;
                    if (!pLine.Closed)
                    {
                        using (Circle circ = new Circle())
                        {
                            circ.Center = pLine.EndPoint;
                            circ.Radius = .5;
                            circ.ColorIndex = 0;

                            btr.AppendEntity(circ);
                            tr.AddNewlyCreatedDBObject(circ, true);
                            ed.WriteMessage("\nError: Open polyline on part " + blkName);
                            partsErrorLog.AppendLine("Warning: Open polyline on part " + blkName);
                        }
                        valid = false;
                    }
                    //check for overlapping polylines
                    foreach (ObjectId secondId in partCol)
                    {
                        if (secondId.ObjectClass.DxfName.ToUpper() == "LWPOLYLINE" &&
                            secondId != msId)
                        {
                            Polyline plineTwo = tr.GetObject(secondId, OpenMode.ForRead) as Polyline;
                            if (pLine.StartPoint == plineTwo.StartPoint &&
                                pLine.EndPoint == plineTwo.EndPoint)
                            {
                                using (Circle circ = new Circle())
                                {
                                    circ.Center = pLine.EndPoint;
                                    circ.Radius = .5;
                                    circ.ColorIndex = 0;

                                    btr.AppendEntity(circ);
                                    tr.AddNewlyCreatedDBObject(circ, true);
                                }
                                //LIST NAME OF OFFENDING PART NUMBER
                                ed.WriteMessage("\nError: Duplicate polyline on part " + blkName);
                                partsErrorLog.AppendLine("Warning: Duplicate polyline on part " + blkName);
                                valid = false;
                            }
                        }
                    }
                }
                else if (msId.ObjectClass.DxfName.ToUpper() == "POLYLINE")
                {
                    valid = false;
                    ed.WriteMessage("\nError: Using 3D Polyline in part " + blkName );
                    partsErrorLog.AppendLine("Critical: Using 3D Polyline in part " + blkName);
                }

                //check if the entity is on an invalid layer
                Entity ent = tr.GetObject(msId, OpenMode.ForRead) as Entity;
                //BlockReference blkRef = tr.GetObject(msId, OpenMode.ForRead) as BlockReference;
                if (ent.Layer.ToUpper() == "DRILL CIRCLE" ||
                    ent.Layer.ToUpper() == "DRILL NODE")
                {
                    valid = false;
                    ed.WriteMessage("\nError: Using old drill layer");
                    partsErrorLog.AppendLine("Warning: Using old drill layer in part " + blkName);
                }
            }         
            return valid;
        }

        //search db for a block of a certain name, deletes everything in that block including the block itself
        public void cleanSlate(Transaction tr, BlockTable bt, BlockTableRecord btr, Editor ed)
        {
            foreach(ObjectId id in btr)
            {
                if(id.ObjectClass.DxfName.ToUpper() == "INSERT")
                {
                    BlockReference blkRef = tr.GetObject(id, OpenMode.ForWrite) as BlockReference;
                    string blockName = "";
                    //if there are any undefined blocks around, then they wont have names to look for
                    try
                    { blockName = blkRef.Name; }
                    catch (System.Exception)
                    { }
                    if (blockName == "NestedPartsContainer")
                    {                       
                        Extents3d window = (Extents3d)blkRef.Bounds;

                        //use the bounds of the block to select everything inside
                        SelectionSet blocks = null;
                        PromptSelectionResult psr = ed.SelectCrossingWindow(window.MinPoint, window.MaxPoint);
                        if(psr.Status == PromptStatus.OK)
                            {blocks = psr.Value;}
                        
                        //start a list for deltetion including the container
                        ObjectIdCollection deleteMe = new ObjectIdCollection(blocks.GetObjectIds());
                        deleteMe.Add(id);

                        //loop through everything selected
                        //if its a block then delete it                       
                        foreach(ObjectId objId in deleteMe)
                        {
                            if(objId.ObjectClass.DxfName.ToUpper() == "INSERT")
                            {
                                BlockReference partBlock = tr.GetObject(objId, OpenMode.ForWrite) as BlockReference;
                                partBlock.Erase();
                            }
                        }
                    }
                }
            }
        }

        //creates a block around the blocked parts
        public void surroundParts(Transaction tr,BlockTable bt, BlockTableRecord btr,Database db,
            Point3d extOld, double nowMaxX,double YoffSet)
        {
            Point3d extNowMin = db.Extmin;

            ObjectId blkRecId = ObjectId.Null;
            if(!bt.Has("NestedPartsContainer"))
            {
                using(BlockTableRecord btrInsert = new BlockTableRecord())  
                {
                    btrInsert.Name = "NestedPartsContainer";
                    using(Polyline frame = new Polyline())  
                    {
                        frame.AddVertexAt(0, new Point2d(extOld.X, extOld.Y), 0, 0, 0);
                        frame.AddVertexAt(1, new Point2d(nowMaxX, extOld.Y), 0, 0, 0);
                        frame.AddVertexAt(2, new Point2d(nowMaxX, extOld.Y - YoffSet), 0, 0, 0);
                        frame.AddVertexAt(3, new Point2d(extOld.X, extOld.Y - YoffSet), 0, 0, 0);
                        frame.Closed = true;
                        frame.Layer = "0";

                        //add the object to the temporary btr
                        btrInsert.AppendEntity(frame);

                        //add the object to the bt & transaction
                        bt.UpgradeOpen(); //get write access
                        bt.Add(btrInsert);
                        tr.AddNewlyCreatedDBObject(btrInsert, true);
                    }
                }
                blkRecId = bt["NestedPartsContainer"];

                if(blkRecId != ObjectId.Null)
                {
                    using (BlockReference blkRef = new BlockReference(new Point3d(0,0,0),blkRecId))
                    {
                        blkRef.Layer = "0";
                        btr.AppendEntity(blkRef);
                        tr.AddNewlyCreatedDBObject(blkRef, true);
                    }
                }
            }
            
            
        }

        //clones a collection of objects into a new collection
        public ObjectIdCollection cloneCollection(ObjectIdCollection partCol, Transaction tr, BlockTableRecord btr)
        {
            ObjectIdCollection clonedCol = new ObjectIdCollection();

            foreach(ObjectId id in partCol)
            {
                Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                Entity copyEnt = ent.Clone() as Entity;
                //add cloned object to the BTR and DB
                btr.AppendEntity(copyEnt);
                tr.AddNewlyCreatedDBObject(copyEnt, true);
                clonedCol.Add(copyEnt.Id);
            }

            return clonedCol;
        }

        //update blocks to anthing in its frame
        private void reblock(string blkName, BlockTable bt, Transaction tr, Database db, 
            ObjectIdCollection objIdCol, Extents3d window, 
            double ydist, double varX, double varY, double quantity,
            Extents3d extParts)
        {
            ObjectId blkRecId = ObjectId.Null;
            //check for a block with the same name
            //if it doesnt exist, then create the block name

            //if the old block is floating around, as a copy somewhere we need to delte the blkref
            // Once deleted from the block table the block stil exists on screen but is unreferenced
            if(bt.Has(blkName))
            {
                BlockTableRecord oldBlock = tr.GetObject(bt[blkName], OpenMode.ForWrite) as BlockTableRecord;
                oldBlock.Erase();
            }            

            //now that the old block doesnt exist or never existed, lets make it
            if (!bt.Has(blkName))
            {
                using (BlockTableRecord newBtr = new BlockTableRecord())
                {
                    //set the name of the block reference
                    newBtr.Name = blkName;
                    //set the origin to the upper left part of the block
                    newBtr.Origin = new Point3d(extParts.MinPoint.X, extParts.MaxPoint.Y, 0);

                    //move access rights from read to write
                    bt.UpgradeOpen();
                    //ad to the block table record
                    bt.Add(newBtr);
                    newBtr.AssumeOwnershipOf(objIdCol);
                    tr.AddNewlyCreatedDBObject(newBtr, true);
                }
                blkRecId = bt[blkName];
            }

            //insert the block into the current space
            if (blkRecId != ObjectId.Null)
            {
                for (int i = 0; i < quantity; i++)
                {
                    Point3d insertPoint = new Point3d(varX, varY, 0);
                    using (BlockReference partRef = new BlockReference(insertPoint, blkRecId))
                    {
                        BlockTableRecord btr = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        btr.AppendEntity(partRef);
                        tr.AddNewlyCreatedDBObject(partRef, true);
                    }
                    varY = varY - (ydist + 10);
                }
            } 
        }

        //clear out unreference blocks
        public void clearUnrefedBlocks(Document doc, Editor ed, Database db)
        {
            using(Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForWrite) as BlockTable;
                //loop through all objects
                foreach(ObjectId objId in bt)
                {
                    //if the block doesnt have any content and isnt a paperspace layout, its likely an undefined block
                    BlockTableRecord btr = tr.GetObject(objId, OpenMode.ForWrite) as BlockTableRecord;
                    if (btr.GetBlockReferenceIds(false, false).Count == 0 && !btr.IsLayout)
                        btr.Erase();
                }
                tr.Commit();
            }
        }

        //pass a collection of Ids, turn into a block
        private void blocking(string name, ObjectIdCollection objIdCol, BlockTable bt, Transaction tr, Database db, BlockTableRecord btr)
        {
            ObjectId blkRecId = ObjectId.Null;
            if (!bt.Has(name))
            {
                //if not, lets create it (it really shouldn't exist yet)
                using (BlockTableRecord btrInsert = new BlockTableRecord())
                {
                    btrInsert.Name = name;

                    bt.UpgradeOpen();
                    bt.Add(btrInsert);
                    btrInsert.AssumeOwnershipOf(objIdCol);
                    tr.AddNewlyCreatedDBObject(btrInsert, true);
                }
                blkRecId = bt[name];
            }
            if (blkRecId != ObjectId.Null)
            {
                //recreate the btr
                BlockTableRecord btrInsert = tr.GetObject(blkRecId, OpenMode.ForRead) as BlockTableRecord;
                using (BlockReference blkRef = new BlockReference(new Point3d(0, 0, 0), blkRecId))
                {
                    btr.AppendEntity(blkRef);
                    tr.AddNewlyCreatedDBObject(blkRef, true);
                }
            }
        }

        //another strategy to updating the blocks
        private void blockbuster(String blkName, BlockTable bt, Transaction tr, Database db, ObjectIdCollection objIdCol, Extents3d window)
        {
            ObjectId blkRecId = ObjectId.Null;
            //check for a block with the same name
            //if it doesnt exist, then create the block name
            if (!bt.Has(blkName))
            {
                using (BlockTableRecord newBtr = new BlockTableRecord())
                {
                    //set the name of the block reference
                    newBtr.Name = blkName;
                    //set the origin
                    newBtr.Origin = window.MinPoint;

                    //move access rights from read to write
                    bt.UpgradeOpen();
                    //ad to the block table record
                    bt.Add(newBtr);
                    newBtr.AssumeOwnershipOf(objIdCol);
                    tr.AddNewlyCreatedDBObject(newBtr, true);

                    //Application.ShowAlertDialog("New block generated.");
                }
                blkRecId = bt[blkName];
            }
            else
            {
                using(BlockReference oldBlock = new BlockReference(new Point3d(0,0,0), blkRecId))
                {
                    BlockTableRecord oldBTR = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                    oldBTR.AppendEntity(oldBlock);
                    tr.AddNewlyCreatedDBObject(oldBlock, true);

                    using(DBObjectCollection oldStuff = new DBObjectCollection())
                    {
                        oldBlock.Explode(oldStuff);
                    }
                }
            }

            //insert the block into the current space
            if (blkRecId != ObjectId.Null)
            {
                using (BlockReference partRef = new BlockReference(window.MinPoint, blkRecId))
                {
                    BlockTableRecord btr = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                    btr.AppendEntity(partRef);
                    tr.AddNewlyCreatedDBObject(partRef, true);
                }
            }
        }

        //attempts to purge the purgable data
        private static bool purgeAll(Database db, Transaction tr, Editor ed, BlockTable bt)
        {
            bool contPurge = true;
            ObjectIdCollection objIdCol = new ObjectIdCollection();
            foreach (ObjectId objId in bt)
            { objIdCol.Add(objId); }

            //removes things that cannot be purged from the purge list
            db.Purge(objIdCol);
            foreach (ObjectId obj in objIdCol)
            {
                DBObject dbObj = tr.GetObject(obj, OpenMode.ForWrite) as DBObject;
                try
                { dbObj.Erase(); }
                catch (Autodesk.AutoCAD.Runtime.Exception ex)
                { Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("Error:\n" + ex.Message); }
            }

            foreach (ObjectId objId in bt)
            { objIdCol.Add(objId); }
            db.Purge(objIdCol);
            if (objIdCol.Count == 0)
            { contPurge = false; }

            return contPurge;
        }

        //loops through a collection to explode block references
        private static ObjectIdCollection explodeBlock(Transaction tr, ObjectIdCollection objIdCol, BlockTableRecord btr, bool contExplode = false)
        {            
            foreach(ObjectId obj in objIdCol)
            {
                Entity ent = null;
                try
                { ent = tr.GetObject(obj, OpenMode.ForWrite) as Entity; }
                catch (Autodesk.AutoCAD.Runtime.Exception e)//If an ent doesnt have an ID registered it will throw an exception
                { }
                    
                if (obj.ObjectClass.DxfName.ToUpper() == "INSERT")
                {
                    DBObjectCollection tempHolder = new DBObjectCollection();
                    ent.Explode(tempHolder);

                    //tell it to run again, until nothing returns as a block
                    contExplode = true;

                    //remove the block from our list
                    objIdCol.Remove(obj);

                    //add the exploded parts to the objectId Collection
                    foreach(Entity dbObj in tempHolder)
                    {
                        //add the exploded part to the modelspace
                        btr.AppendEntity(dbObj);
                        tr.AddNewlyCreatedDBObject(dbObj,true);
                        //add the part to the list
                        objIdCol.Add(dbObj.Id);
                    }
                }
            }

            //loop until the explosion has gone deep enough
            if (contExplode)
            { objIdCol = explodeBlock(tr, objIdCol, btr); }           

            return objIdCol;
        }

        //current naming scheme being used is the group#, item#, part#, partname
        public static string getName(BlockReference blkRef, Transaction tr, StringBuilder group, StringBuilder style)
        {
            AttributeCollection atts = blkRef.AttributeCollection;
            string[] nameParts = new string[4];

            //iterate through the attributes to find the parts for a name
            foreach(ObjectId objId in atts)
            {
                DBObject dbObj = tr.GetObject(objId, OpenMode.ForRead) as DBObject;
                AttributeReference attRef = dbObj as AttributeReference;
                

                if(attRef.Tag.Contains("SUITE"))
                {nameParts[0]=attRef.TextString.Trim();}
                else if(attRef.Tag.Contains("ITEM"))
                {nameParts[1]=attRef.TextString.Trim();}
                else if (attRef.Tag.Contains("PART-NO"))
                { nameParts[2] = attRef.TextString.Trim(); }
                else if (attRef.Tag.Contains("PARTNAME"))
                { nameParts[3] = attRef.TextString.Trim(); }
            }

            //might add the version to the end
            string name = String.Format("{0}-{1}-{2}-{3}",
                nameParts[0], nameParts[1], nameParts[2], nameParts[3]);

            //compare expected numbers to those found in file
            //if comparison yeilds something, add to error log
            if(group.ToString() != nameParts[0] || style.ToString() != nameParts[1])
            {
                partsErrorLog.AppendLine("Warning: Filename " + name + " has unexpected naming");
            }
            return name;
        }

        //looks through the blkref to find the quantity of parts to list
        public static double getQuantity(BlockReference blkRef, Transaction tr)
        {
            AttributeCollection atts = blkRef.AttributeCollection;
            double quantity = 0;

            //iterates through the attributes to find the tag we want
            foreach(ObjectId objId in atts)
            {
                DBObject dbObj = tr.GetObject(objId, OpenMode.ForRead) as DBObject;
                AttributeReference attRef = dbObj as AttributeReference;

                if(attRef.Tag.Contains("QUANTITY"))
                { 
                    var holder = attRef.TextString;
                    quantity = Convert.ToDouble(holder);
                }
            }
            return quantity;
        }

        //method to find the extents of a collection
        public Extents3d getExtents(Transaction tr, ObjectIdCollection ids)
        {
            var ext = new Extents3d();
            foreach(ObjectId id in ids)
            {
                Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                if(ent != null)
                { ext.AddExtents(ent.GeometricExtents); }
            }
            return ext;
        }

        //output the message box with all collected warnings
        private static void warnBox()
        {
            if (partsErrorLog.ToString() == null || partsErrorLog.ToString() == "")
            {
                partsErrorLog.Append("No immdediate errors found");
            }
            //message box shows errors or lack therof
            //message box to display error messages found during operation            
            new Thread(new ThreadStart(delegate
            {
                MessageBox.Show
                (
                  partsErrorLog.ToString(),
                  "Parts Error Check:",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Warning
                );
            })).Start();
        }

        //set up columns for datatable
        public static System.Data.DataTable createTableHeader()
        {
            // Set up table header
            System.Data.DataTable table = new System.Data.DataTable("Parts list");

            //list table column headers
            System.Data.DataColumn[] cols ={
                //Block name, partname, part quantity, spec block id used, scale of spec
                                  new System.Data.DataColumn ("BlockName", typeof(string)),
                                  new System.Data.DataColumn ("PartName", typeof(string)),
                                  new System.Data.DataColumn ("Quantity", typeof(double)),
                                  new System.Data.DataColumn ("BlockId",typeof(Handle)),
                                  new System.Data.DataColumn ("BlockIdPart",typeof(Handle)),
                                  new System.Data.DataColumn ("Scale",typeof(double))
                              };
            //load column headers into datatable
            table.Columns.AddRange(cols);
            return table;
        }

        //moves frame contents to new location following simple rules to arrange based on scale
        private static Point3d moveSpec(Database db, DataRow row, Point3d lastPoint, bool newRow = false)
        {
            //might have conversion issue of blockID saved as string
            var specId = row["BlockId"];
            var partId = row["BlockIdPart"];
            ObjectId specID = objectIdFromHandle(specId.ToString(), db);
            ObjectId part = objectIdFromHandle(partId.ToString(), db);

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockReference specRef = tr.GetObject(specID, OpenMode.ForWrite) as BlockReference;
                BlockReference partRef = tr.GetObject(part, OpenMode.ForWrite) as BlockReference;

                //get dimensions of the spec
                Extents3d window = (Extents3d)specRef.Bounds;

                Vector3d movement = new Vector3d();
                //take the last point and move to being lined up with first column but down the height of the current spec +2
                if (newRow)
                {
                    movement = new Point3d(window.MinPoint.X, window.MaxPoint.Y, 0).
                        GetVectorTo(new Point3d(lastPoint.X, lastPoint.Y - 2, lastPoint.Z));
                }
                else //take last point and move to being lined up with previous row but to right by width of frame +2
                {
                    movement = new Point3d(window.MinPoint.X, window.MinPoint.Y, 0).
                        GetVectorTo(new Point3d(lastPoint.X + 2, lastPoint.Y, lastPoint.Z));
                }

                specRef.TransformBy(Matrix3d.Displacement(movement));
                partRef.TransformBy(Matrix3d.Displacement(movement));
                tr.Commit();

                window = (Extents3d)specRef.Bounds;
                lastPoint = new Point3d(window.MaxPoint.X, window.MinPoint.Y, 0);
            }
            return lastPoint;
        }

        //using handle stored as a string acquires blockID
        private static ObjectId objectIdFromHandle(string stringId, Database db)
        {
            Int64 idHex = Int64.Parse(stringId.ToString(), NumberStyles.AllowHexSpecifier);
            ObjectId specID = db.GetObjectId(false, new Handle(idHex), 0);
            return specID;
        }
    }
}
