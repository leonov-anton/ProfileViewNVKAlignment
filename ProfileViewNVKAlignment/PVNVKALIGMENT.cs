using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Geometry;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using System;

namespace ProfileViewNVKAlignment
{
    public class PVNVKALIGMENT
    {
        [CommandMethod("PVNVKALIG")]
        static public void PipesAlignment()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            PromptEntityOptions peo = new PromptEntityOptions("\nВыберите главную трубу:");
            peo.SetRejectMessage("\nВыбраный объект не труба");
            peo.AddAllowedClass(typeof(ProfileViewPart), true);
            PromptEntityResult per = ed.GetEntity(peo);
            if(per.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nВыполнение прервано");
                return;
            }

            PromptSelectionOptions pso = new PromptSelectionOptions();
            pso.MessageForAdding = "\nВыберите выравниваемые трубы:";
            PromptSelectionResult psr = ed.GetSelection(pso);
            if (psr.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nВыполнение прервано");
                return;
            }

            PromptKeywordOptions pko = new PromptKeywordOptions("\nСохранить уклон или изменить по главной трубе:");
            pko.Keywords.Add("Сохранить");
            pko.Keywords.Add("По главной");
            pko.Keywords.Default = "По главной";
            PromptResult des = ed.GetKeywords(pko);
            if (des.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nВыполнение прервано");
                return;
            }

            using (Transaction t = doc.TransactionManager.StartTransaction())
            {
                ProfileViewPart mainPvpObj = (ProfileViewPart)t.GetObject(per.ObjectId, OpenMode.ForRead);
                Pipe mainPipe = (Pipe)t.GetObject(mainPvpObj.ModelPartId, OpenMode.ForWrite);

                DBObjectCollection secPipesObjs = new DBObjectCollection();

                foreach (ObjectId pvpId in psr.Value.GetObjectIds())
                {
                    if (pvpId.ObjectClass.Name == "AeccDbGraphProfileNetworkPart")
                    {
                        ProfileViewPart secPVPObj = (ProfileViewPart)t.GetObject(pvpId, OpenMode.ForRead);
                        if (secPVPObj.ModelPartId.ObjectClass.Name == "AeccDbPipe")
                        {
                            Pipe pipeObj = (Pipe)t.GetObject(secPVPObj.ModelPartId, OpenMode.ForRead);
                            double diamPVP = Math.Round(secPVPObj.Bounds.Value.MaxPoint.X - secPVPObj.Bounds.Value.MinPoint.X, 3);
                            double diamPipe = Math.Round(pipeObj.WallThickness * 2 + pipeObj.InnerDiameterOrWidth, 3);
                            if (pipeObj.NetworkId == mainPipe.NetworkId 
                                && pipeObj is Pipe 
                                && secPVPObj.ModelPartId != mainPvpObj.ModelPartId 
                                && diamPVP != diamPipe)
                                secPipesObjs.Add(pipeObj);
                        }
                    }
                }

                Pipe privPipe = mainPipe;
                double offset;
                double slope;
                string holdPoint;
                Pipe nextPipe;

                while (secPipesObjs.Count > 0)
                {
                    foreach (Pipe pipe in secPipesObjs)
                    {
                        nextPipe = pipe;
                        // присоединение началом трубы к концу предыдущей трубы
                        if (privPipe.EndStructureId != null && privPipe.EndStructureId == pipe.StartStructureId)
                        {
                            offset = GetOffset(privPipe.InnerDiameterOrWidth, pipe.InnerDiameterOrWidth, privPipe.EndPoint.Z, pipe.StartPoint.Z);
                            slope = -privPipe.Slope;
                            holdPoint =  "sp";
                            goto Found;
                        }
                        // присоединение концом трубы к концу предыдущей трубы
                        else if (privPipe.EndStructureId != null && privPipe.EndStructureId == pipe.EndStructureId)
                        {
                            offset = GetOffset(privPipe.InnerDiameterOrWidth, pipe.InnerDiameterOrWidth, privPipe.EndPoint.Z, pipe.EndPoint.Z);
                            slope = -privPipe.Slope;
                            holdPoint = "ep";
                            goto Found;
                        }
                        // присоединение началом трубы к началу предыдущей трубы
                        else if (privPipe.StartStructureId != null && privPipe.StartStructureId == pipe.StartStructureId)
                        {
                            offset = GetOffset(privPipe.InnerDiameterOrWidth, pipe.InnerDiameterOrWidth, privPipe.StartPoint.Z, pipe.StartPoint.Z);
                            slope = privPipe.Slope;
                            holdPoint = "sp";
                            goto Found;
                        }
                        // присоединение концом трубы к началу предыдущей трубы
                        else if (privPipe.StartStructureId != null && privPipe.StartStructureId == pipe.EndStructureId)
                        {
                            offset = GetOffset(privPipe.InnerDiameterOrWidth, pipe.InnerDiameterOrWidth, privPipe.StartPoint.Z, pipe.EndPoint.Z);
                            slope = privPipe.Slope;
                            holdPoint = "ep";
                            goto Found;
                        }
                    }

                    ed.WriteMessage("\nТрубы не соединены. Выполнение прервано.");
                    t.Commit();
                    return;

                    Found:
                        SetNewElevation(db, nextPipe, slope, offset, des.StringResult, holdPoint);
                        privPipe = nextPipe;
                        secPipesObjs.Remove(nextPipe);
                }
                t.Commit();
            }
        }

        static double GetOffset (double privPipeInnDim, double pipeInnDim, double privPipeZ, double pipeZ)
        {
            double offset;

            if (privPipeInnDim >= pipeInnDim)
                offset = privPipeZ - privPipeInnDim / 2 - pipeZ + pipeInnDim / 2;
            else
                offset = privPipeZ + privPipeInnDim / 2 - pipeZ - pipeInnDim / 2;

            return offset;
        }

        static void SetNewElevation (Database db, Pipe pipe, double slope, double offset, string choice, string holdPoint)
        {
            using (Transaction tmpt = db.TransactionManager.StartTransaction())
            {
                pipe.UpgradeOpen();
                if (choice == "Сохранить")
                {
                    pipe.StartPoint = new Point3d(pipe.StartPoint.X, pipe.StartPoint.Y, pipe.StartPoint.Z + offset);
                    pipe.EndPoint = new Point3d(pipe.EndPoint.X, pipe.EndPoint.Y, pipe.EndPoint.Z + offset);
                }
                else
                {
                    switch (holdPoint)
                    {
                        case "sp":
                            pipe.StartPoint = new Point3d(pipe.StartPoint.X, pipe.StartPoint.Y, pipe.StartPoint.Z + offset);
                            pipe.SetSlopeHoldStart(slope);
                            break;
                        case "ep":
                            pipe.EndPoint = new Point3d(pipe.EndPoint.X, pipe.EndPoint.Y, pipe.EndPoint.Z + offset);
                            pipe.SetSlopeHoldEnd(slope);
                            break;
                    }
                }
                Structure startStruct = (Structure)tmpt.GetObject(pipe.StartStructureId, OpenMode.ForWrite);
                Structure endStruct = (Structure)tmpt.GetObject(pipe.EndStructureId, OpenMode.ForWrite);

                startStruct.ResizeByPipeDepths();
                endStruct.ResizeByPipeDepths();

                tmpt.Commit();
            }
        }
    }
}
