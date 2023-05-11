using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace Blocks_for_Nesting
{
    class Insert
    {
        public static void insertSpec(BlockReference part, DataRow partInfo)
        {
            //!!!!!!!!!!!!!!!!!!!!! HARDCODED INFO
            //vvvvvvvvvvvvvvvvv Spec sheet location in code
            string specSheet = @"Y:\Engineering\Drawings\Blocks\Standard Forms\SF_CNC_BORDER_V2.dwg";
            List<double> insertArea = new List<double>();
            insertArea.Add(10.37497925);
            insertArea.Add(6.49834463);
            Point3d targetCenter = new Point3d(5.18748962, 4.09514693, 0);
            //-----------------------------------------------------------------------

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            //get coords of the part
            Extents3d window = (Extents3d)part.Bounds;
            //find scale to fit min|max of selection into insert area
            double scale = calcScale(window.MinPoint, window.MaxPoint, insertArea);
            //find insertpoint using scale to offset
            Point3d insertPoint = getInsertPoint(window.MinPoint, window.MaxPoint, scale, targetCenter);
            //find expected styleID and group number from filename
            #region find name
            string fileName = Path.GetFileNameWithoutExtension(doc.Name);
            StringBuilder group = new StringBuilder();
            StringBuilder style = new StringBuilder();
            bool secondPart = false;
            bool firstPart = false;
            char[] groupList = { ' ', '-', 'A', 'C' };
            char[] styleList = { ' ', 'S', 'N' };//might need to change this to be more general
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

            //insertblock wrap in another transaction
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockInserter(specSheet, insertPoint, scale, group.ToString(), style.ToString(), partInfo);
                tr.Commit();
            }
        }

        private static void BlockInserter(string specSheet, Point3d insertPoint, double scale, string group, string style, DataRow partInfo)
        {
            string specName = Path.GetFileNameWithoutExtension(specSheet);
            Database dbCurrent = Application.DocumentManager.MdiActiveDocument.Database;

            using (Transaction trCurrent = dbCurrent.TransactionManager.StartTransaction())
            {
                //open block table for read
                BlockTable btCurrent = trCurrent.GetObject(dbCurrent.BlockTableId, OpenMode.ForRead) as BlockTable;

                //check if spec is already loaded into drawing
                ObjectId blkRecId = ObjectId.Null;
                if (!btCurrent.Has(specName))
                {
                    //open db to other file
                    Database db = new Database(false, true);
                    try
                    { db.ReadDwgFile(specSheet, System.IO.FileShare.Read, false, ""); }
                    catch (System.Exception)
                    {
                        return;
                    }
                    dbCurrent.Insert(specName, db, true);

                    blkRecId = btCurrent[specName];
                }
                else
                { blkRecId = btCurrent[specName]; }

                //now insert block into current space
                if (blkRecId != ObjectId.Null)
                {
                    //create btr for the inserted block
                    BlockTableRecord btrInsert = trCurrent.GetObject(blkRecId, OpenMode.ForRead) as BlockTableRecord;
                    using (BlockReference blkRef = new BlockReference(insertPoint, blkRecId))
                    {
                        BlockTableRecord btrCurrent = trCurrent.GetObject(dbCurrent.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        //scale the frame using insert point and scalefactor
                        blkRef.TransformBy(Matrix3d.Scaling(scale, insertPoint));

                        //add the frame to the btr
                        btrCurrent.AppendEntity(blkRef);
                        trCurrent.AddNewlyCreatedDBObject(blkRef, true);

                        //update info into table
                        partInfo["Scale"] = scale;
                        //get specific blockref blockID  <<<<<<<<<<<<<<<<<<<<<
                        partInfo["BlockId"] = blkRef.Handle;

                        // Verify block table record has attribute definitions associated with it
                        if (btrInsert.HasAttributeDefinitions)
                        {
                            // Add attributes from the block table record
                            foreach (ObjectId objID in btrInsert)
                            {
                                DBObject dbObj = trCurrent.GetObject(objID, OpenMode.ForRead) as DBObject;
                                if (dbObj is AttributeDefinition)
                                {
                                    AttributeDefinition acAtt = dbObj as AttributeDefinition;
                                    if (!acAtt.Constant)
                                    {
                                        using (AttributeReference acAttRef = new AttributeReference())
                                        {
                                            acAttRef.SetAttributeFromBlock(acAtt, blkRef.BlockTransform);
                                            acAttRef.Position = acAtt.Position.TransformBy(blkRef.BlockTransform);

                                            acAttRef.TextString = acAtt.TextString;

                                            blkRef.AttributeCollection.AppendAttribute(acAttRef);
                                            trCurrent.AddNewlyCreatedDBObject(acAttRef, true);
                                        }
                                    }
                                }
                            }
                            // Write new data into the block
                            //
                            AttributeCollection attCol = blkRef.AttributeCollection;
                            foreach (ObjectId objID in attCol)
                            {
                                DBObject dbObj = trCurrent.GetObject(objID, OpenMode.ForRead) as DBObject;
                                AttributeReference acAttRef = dbObj as AttributeReference;
                                //initials need to be in a specific file location or registry loc
                                if (acAttRef.Tag.Contains("SUITE"))
                                { acAttRef.TextString = group; }
                                else if (acAttRef.Tag.Contains("ITEM"))
                                { acAttRef.TextString = style; }
                                else if (acAttRef.Tag.Contains("PART-NO"))
                                { }
                                else if (acAttRef.Tag.Contains("PARTNAME"))
                                { acAttRef.TextString = partInfo["PartName"].ToString(); }
                                else if (acAttRef.Tag.Contains("QUANTITY"))
                                { acAttRef.TextString = partInfo["Quantity"].ToString(); }
                                else if (acAttRef.Tag.Contains("THICK"))
                                { }
                                else if (acAttRef.Tag.Contains("SCALE"))
                                { acAttRef.TextString = blkRef.ScaleFactors.X.ToString() + "x"; }
                            }
                        }
                    }
                    trCurrent.Commit();
                }
            }
        }

        //Calculate scale based on assemed dimensions of spec and the dimensions of area to be surrounded
        private static double calcScale(Point3d minPt, Point3d maxPt, List<double> insertArea)
        {
            double distX = Math.Abs(maxPt.X - minPt.X);
            double distY = Math.Abs(maxPt.Y - minPt.Y);
            //add on 10% for some margin
            distX = distX * 1.1;
            distY = distY * 1.1;

            //get dimensions of insert area of spec sheet
            double frameDistX = insertArea[0];
            double frameDistY = insertArea[1];

            double scaleX = distX / frameDistX;
            double scaleY = distY / frameDistY;

            //whichever is higher is the scale to work from
            int scale;
            if (scaleX > scaleY)
                scale = Convert.ToInt16(scaleX);
            else
                scale = Convert.ToInt16(scaleY);

            //after converting to int, decimal places were lost, adding now to compensate for that
            scale++;

            return Convert.ToDouble(scale);
        }

        //assumes that the default insertPoint on the insertspec is 0,0,0****
        private static Point3d getInsertPoint(Point3d minPt, Point3d maxPt, double scale, Point3d targetCenter)
        {
            //calc the center of the selected area
            Point3d selectionCenter = getCentroid(minPt, maxPt);

            //using the distance of the center of the specsheet insert area from its ref point (0,0,0)
            //multiply said distance by the scale factor
            //use this new distance to offset from our selection's center for the insert point
            double x = selectionCenter.X - (targetCenter.X * scale);
            double y = selectionCenter.Y - (targetCenter.Y * scale);

            return new Point3d(x, y, 0);
        }

        private static Point3d getCentroid(Point3d minPt, Point3d maxPt)
        {
            //formula
            //total distance, divided by two, added to original point
            double distX = (Math.Abs(maxPt.X - minPt.X) / 2) + minPt.X;
            double distY = (Math.Abs(maxPt.Y - minPt.Y) / 2) + minPt.Y;
            //could do the same for Z, but probaly not necissary

            return new Point3d(distX, distY, 0);
        }
    }
}
