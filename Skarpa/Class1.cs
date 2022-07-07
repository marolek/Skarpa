using System;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.GraphicsInterface;
using Autodesk.AutoCAD.Geometry;
using Autodesk.Civil.ApplicationServices;


namespace Skarpa
{
    public class Class1 : IExtensionApplication
    {
        Double odstep = 0.5;

        private TransientManager tm = TransientManager.CurrentTransientManager;
        private Line linia;
        public Curve gKraw;

        public void Initialize()
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            acDoc.Editor.WriteMessage("\nWczytano dodatek Skarpa 2023");
        }

        public void Terminate()
        {
            throw new System.Exception("The method or operation is not implemented.");
        }

        public void ZmienOdstep()
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            PromptDoubleOptions odstepOpts;
            odstepOpts = new PromptDoubleOptions("\nPodaj skalę kreskowania: ");

            odstepOpts.AllowNone = false;
            odstepOpts.DefaultValue = odstep;
            odstepOpts.AllowZero = false;
            odstepOpts.UseDefaultValue = true;

            PromptDoubleResult odstepRes = ed.GetDouble(odstepOpts);
            if (odstepRes.Status != PromptStatus.OK) {return;}
            odstep = odstepRes.Value;
        }

        public void TrackMouse(Object sender, PointMonitorEventArgs e)
        {
            Point3d mousePoint = e.Context.RawPoint;
            Point3d pol = gKraw.GetClosestPointTo(mousePoint, false);

            linia.StartPoint = mousePoint;
            linia.EndPoint = pol;
            try
            {
                double station = gKraw.GetDistAtPoint(gKraw.GetClosestPointTo(pol, false));
                e.AppendToolTipText("Pikieta: " + station.ToString("N2"));
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                String blad = String.Format("\nJakiś błąd w TrackMouse: {0}",ex.Message);
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(blad);
            }
            tm.UpdateTransient(linia, new IntegerCollection());
        }

        public double Rysuj_linie(Point3d pktStart, Point3d pktKoniec)
        {
            Document oDWG = Application.DocumentManager.MdiActiveDocument;
            Database oDB = oDWG.Database;
            Autodesk.AutoCAD.DatabaseServices.Transaction oTrans = oDB.TransactionManager.StartTransaction();
            BlockTable oBT = oTrans.GetObject(oDB.BlockTableId, OpenMode.ForRead) as BlockTable;
            BlockTableRecord oBTR = oTrans.GetObject(oBT[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
            double retVal = 0;

            try
            {
                Line myLine = new Line(pktStart, pktKoniec);

                oBTR.AppendEntity(myLine);
                oTrans.AddNewlyCreatedDBObject(myLine, true);
                oTrans.Commit();
                retVal = myLine.Length;
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                String blad = String.Format("\nBłąd rysuj linie: {0}",ex.Message);
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(blad);
            }
            finally
            {
                oTrans.Dispose();
            }

            return retVal;
        }

        [Autodesk.AutoCAD.Runtime.CommandMethod("Skarpa2023")] 
        public void HelloWorld()
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            CivilDocument Civdoc = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument;
            Database acCurDb = acDoc.Database;

            PromptNestedEntityOptions XrefEntOpts1;
            XrefEntOpts1 = new PromptNestedEntityOptions("\nWybierz górną krawędź lub <wyjdz>: ");

            PromptNestedEntityOptions XrefEntOpts2;
            XrefEntOpts2 = new PromptNestedEntityOptions("\nWybierz dolną krawędź lub <wyjdz>: ");

            XrefEntOpts1.Keywords.Add("Odstęp");
            XrefEntOpts1.AllowNone = true;

            XrefEntOpts2.Keywords.Add("Odstęp");
            XrefEntOpts2.AllowNone = true;

            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;

            PromptEntityResult entRes1 = ed.GetNestedEntity(XrefEntOpts1);

            if (entRes1.StringResult == "Odstęp") {
                ZmienOdstep();
                return;
            }
            if (entRes1.Status != PromptStatus.OK) {return;}
            
            PromptEntityResult entRes2 = ed.GetNestedEntity(XrefEntOpts2);

            if (entRes2.StringResult == "Odstęp") {
                ZmienOdstep();
                return;
            }
            if (entRes2.Status != PromptStatus.OK) {return;}

            Type[] dozwolone = { typeof(Polyline2d), typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), typeof(Polyline3d), typeof(Circle), typeof(Line), typeof(Arc), typeof(Spline) };

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                DBObject dbObj1 = acTrans.GetObject(entRes1.ObjectId, OpenMode.ForWrite);
                //typ objektu xref
                Type TypXRef = dbObj1.GetType();
                if(!Array.Exists(dozwolone, x => x == TypXRef))
                {
                    ed.WriteMessage("\nGórna krawędź jest nieprawidłowa");
                    return;
                }
                gKraw = (Curve)dbObj1;

                Curve tempGKraw = (Curve)gKraw.Clone();
                tempGKraw.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 4);
                tm.AddTransient(tempGKraw, TransientDrawingMode.Highlight, 128, new IntegerCollection());

                DBObject dbObj2 = acTrans.GetObject(entRes2.ObjectId, OpenMode.ForWrite);

                TypXRef = dbObj2.GetType();
                if(!Array.Exists(dozwolone, x => x == TypXRef))
                {
                    ed.WriteMessage("\nDolna krawędź jest nieprawidłowa");
                    return;
                }
                Curve dKraw = (Curve)dbObj2;

                Curve tempDKraw = (Curve)dKraw.Clone();
                tempDKraw.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 4);
                tm.AddTransient(tempDKraw, TransientDrawingMode.Highlight, 128, new IntegerCollection());

                if (dbObj1.Id == dbObj2.Id)
                {
                    ed.WriteMessage("\nWybrano tą samą polilinię");
                    return;
                }


                Double zakresDOLstart  = dKraw.GetDistAtPoint(dKraw.GetClosestPointTo(gKraw.StartPoint, false));
                Double zakresDOLend = dKraw.GetDistAtPoint(dKraw.GetClosestPointTo(gKraw.EndPoint, false));

                //String s = String.Format("\nZakres dolnej krawędzi od: {0:0.00} do: {1}",zakresDOLstart, zakresDOLend);
                //ed.WriteMessage(s);

                //linia wskazujaca punkt na trasie
                linia = new Line();
                linia.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 1);
                tm.AddTransient(linia, TransientDrawingMode.Main, 128, new IntegerCollection());

                acDoc.Editor.PointMonitor += TrackMouse;

                // wskazanie zakresu
                PromptPointOptions optPointStart = new PromptPointOptions("\nWskaż początek zakresu lub ");
                optPointStart.Keywords.Add("Początek");
                optPointStart.Keywords.Default = "Początek";
                optPointStart.AllowNone = true;

                PromptPointResult resPointStart = acDoc.Editor.GetPoint(optPointStart);

                if (resPointStart.StringResult == "Początek")
                    goto pomin_poczatek;

                if (resPointStart.Status == PromptStatus.OK)
                {
                    try
                    {
                        double pikietaStart = gKraw.GetDistAtPoint(gKraw.GetClosestPointTo(resPointStart.Value, false));
                        acDoc.Editor.WriteMessage("\nPoczątek zakresu: " + pikietaStart.ToString("N2"));
                        zakresDOLstart = dKraw.GetDistAtPoint(dKraw.GetClosestPointTo(resPointStart.Value, false));
                    }
                    // pozycja = pikietaStart
                    catch (Autodesk.AutoCAD.Runtime.Exception ex)
                    {
                        String blad = String.Format("\nBłąd początek zakresu: {0}",ex.Message);
                        ed.WriteMessage(blad);
                    }
                }

            pomin_poczatek:

                PromptPointOptions optPointEnd = new PromptPointOptions("\nWskaż koniec zakresu lub ");
                optPointEnd.Keywords.Add("Koniec");
                optPointEnd.Keywords.Default = "Koniec";
                optPointEnd.AllowNone = true;

                PromptPointResult resPointEnd = acDoc.Editor.GetPoint(optPointEnd);

                if (resPointEnd.StringResult == "Koniec")
                    goto pomin_koniec;

                if (resPointEnd.Status == PromptStatus.OK)
                {
                    try
                    {
                        double pikietaEnd = gKraw.GetDistAtPoint(gKraw.GetClosestPointTo(resPointEnd.Value, false));
                        acDoc.Editor.WriteMessage("\nKoniec zakresu: " + pikietaEnd.ToString("N2"));
                        zakresDOLend = dKraw.GetDistAtPoint(dKraw.GetClosestPointTo(resPointEnd.Value, false));
                    }
                    catch (Autodesk.AutoCAD.Runtime.Exception ex)
                    {
                        String blad = String.Format("\nBłąd koniec zakresu: {0}",ex.Message);
                        ed.WriteMessage(blad);
                    }
                }

            pomin_koniec:

                // jeśli poczatek = koniec (np koło) to oblicz koniec z długości objektu
                if (zakresDOLend == zakresDOLstart)
                {
                    // String info = String.Format("\nOkrąg - początek: {0}, koniec: {1}",zakresDOLstart, zakresDOLend);
                    // ed.WriteMessage(info);
                    double cLen = dKraw.GetDistanceAtParameter(dKraw.EndParam) - dKraw.GetDistanceAtParameter(dKraw.StartParam);
                    zakresDOLend = cLen;
                    // info = String.Format("\nZmienione Okrąg - początek: {0}, koniec: {1}",zakresDOLstart, zakresDOLend);
                    // ed.WriteMessage(info);
                }

                double pozycja = zakresDOLstart; //pozycja kreskowania zaczyna od 'zakresDOLstart'
                
                // gdyby wskazany koniec byl mniejszy niż początek to zamień
                if (zakresDOLend < zakresDOLstart)
                {
                    pozycja = zakresDOLend;
                    zakresDOLend = zakresDOLstart;
                }

                bool polowka = false; //czy rysować krótką kreskę

                while (pozycja < zakresDOLend) // cLen
                {
                    Point3d pktDol = dKraw.GetPointAtDist(pozycja);
                    Point3d pktGora = gKraw.GetClosestPointTo(pktDol, false);

                    Point3d PktDol2D = new Point3d(pktDol.X, pktDol.Y, pktGora.Z); //tymczasowy punkt do obliczenia odległosci między krawędziami (* odstep); Rzędna z punktu górnej krawędzi 
                    double odcinek = PktDol2D.DistanceTo(pktGora) * odstep;

                    if (odcinek < 0.1)
                    {
                        odcinek = 0.1;
                        goto nierysuj;
                    }

                    if (polowka)
                    {
                        if (pktDol.Z > pktGora.Z)
                        {
                            Point3d pktMID = new Point3d(pktDol.X + (pktGora.X - pktDol.X) * odstep, pktDol.Y + (pktGora.Y - pktDol.Y) * odstep, pktDol.Z + (pktGora.Z - pktDol.Z) * odstep);
                            Rysuj_linie(pktMID, pktDol);
                        }
                        else
                        {
                            Point3d pktMID = new Point3d(pktGora.X - (pktGora.X - pktDol.X) * odstep, pktGora.Y - (pktGora.Y - pktDol.Y) * odstep, pktGora.Z - (pktGora.Z - pktDol.Z) * odstep);
                            Rysuj_linie(pktMID, pktGora);
                        }
                    }
                    else
                        Rysuj_linie(pktDol, pktGora);

                nierysuj:
                    pozycja = odcinek + pozycja;
                    polowka = !polowka;
                }

                acDoc.Editor.PointMonitor -= TrackMouse;

                tm.EraseTransient(linia, new IntegerCollection());
                tm.EraseTransient(tempGKraw, new IntegerCollection());
                tm.EraseTransient(tempDKraw, new IntegerCollection());
                acTrans.Commit();
            }
            // ObjectIdCollection alignments = Civdoc.GetAlignmentIds();
        }
    }
}