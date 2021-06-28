//--------------------------------------------------------------------------------------------- 
/// <summary> 
/// Eclipse v16 ESAPI script that exports dose for all daily treatment.
/// The code here is modify from Export3D from
/// https://github.com/VarianAPIs/Varian-Code-Samples/blob/master/Eclipse%20Scripting%20API/plugins/Export3D.cs .
/// </summary> 
/// <license> 
/// 
// Copyright (c) 2016 Varian Medical Systems, Inc. 
// Copyright (c) 2021 44REAM. 
// 
// Permission is hereby granted, free of charge, to any person obtaining a copy  
// of this software and associated documentation files (the "Software"), to deal  
// in the Software without restriction, including without limitation the rights  
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell  
// copies of the Software, and to permit persons to whom the Software is  
// furnished to do so, subject to the following conditions: 
// 
// The above copyright notice and this permission notice shall be included in  
// all copies or substantial portions of the Software. 
// 
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR  
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,  
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL  
// THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER  
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,  
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN  
// THE SOFTWARE. 
/// </license> 
//--------------------------------------------------------------------------------------------- 

using System;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using System.Reflection;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using System.IO;
using System.Windows.Media.Media3D;
using System.Windows.Media;


[assembly: AssemblyVersion("1.0.0.1")]
[assembly: AssemblyFileVersion("1.0.0.1")]
[assembly: AssemblyInformationalVersion("1.0")]



namespace Fraction {
    public class TreatmentFraction {
        public List<PlanSetup> PlanSetupList;
        public int PlanNumber;
        public int NumberOfFraction;

        public TreatmentFraction(List<PlanSetup> planSetupList) {
            PlanSetupList = planSetupList;
            PlanNumber = planSetupList.Count();
            NumberOfFraction = 1;
        }

        public override bool Equals(object obj) => this.Equals(obj as TreatmentFraction);

        public bool Equals(TreatmentFraction p) {
            
            if (p is null) {
                return false;
            }

            if (Object.ReferenceEquals(this, p)) {
                return true;
            }

            if (this.GetType() != p.GetType()) {
                return false;
            }

            if (this.PlanNumber != p.PlanNumber) {
                return false;
            }

            for (int i = 0; i < this.PlanNumber; i++) {
                if (this.PlanSetupList[i].Id != p.PlanSetupList[i].Id) {
                    return false;
                }
            }

            return true;
        }

    }

    public class TreatmentDict {

        public Dictionary<string, string> PlanDict = new Dictionary<string, string>();
        public int PlanNumber = 1;

        public TreatmentDict() {

        }

        public void AddPlanSetup(PlanSetup planSetup) {
            if (!PlanDict.ContainsKey(planSetup.Id)) {
                PlanDict.Add(planSetup.Id, $"dose_{PlanNumber}");
                PlanNumber++;
            }

        }

    }
    class Program {
        private static List<string> PatientIDList = new List<string>();
        private static bool IsExport = true;
        private static string CSVPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\automate_fraction_eso.csv";
        private static string ExportFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\export_fraction";
        [STAThread]
        static void Main(string[] args) {
            try {
                using (Application app = Application.CreateApplication()) {
                    ExportFraction(app, "3334289", "C1", "Lungs");
                    //ExportFractionFromCSV(app);
                }
            } catch (Exception e) {


            }
        }


        static TreatmentDict AddFraction(List<TreatmentFraction> planFraction, Course course) {

            TreatmentDict treatmentDict = new TreatmentDict();
            foreach (TreatmentSession treatmentSession in course.TreatmentSessions) {
                List<PlanSetup> planSetupList = new List<PlanSetup>(); ;

                foreach (PlanTreatmentSession planTreatmentSession in treatmentSession.SessionPlans) {

                    string status = planTreatmentSession.Status.ToString();
                    // check if the plan is derivered
                    // beware CompletedPartially Plan

                    if (status.Equals("Completed")) {
                        PlanSetup planSetup = planTreatmentSession.PlanSetup;

                        // Check if no dose and no image register
                        if (planSetup.Dose == null || planSetup.StructureSet == null ) {
                            planFraction.Clear();
                            throw new Exception("No registered image and no dose data");

                        }
                        treatmentDict.AddPlanSetup(planSetup);
                        planSetupList.Add(planSetup);
                    } else if (status.Equals("CompletedPartially")) {
                        throw new Exception("There are CompletedPartially plan");
                    }

                }

                // check if there are plansetup
                if (planSetupList.Any()) {
                    TreatmentFraction treatmentFraction = new TreatmentFraction(planSetupList);

                    int fractionIndex = planFraction.IndexOf(treatmentFraction);



                    if (fractionIndex == -1) {
                        planFraction.Add(treatmentFraction);
                    } else {
                        planFraction[fractionIndex].NumberOfFraction += 1;
                    }
                }

            }
            return treatmentDict;

        }
        static bool IsCTImageOK(Course course) {

            HashSet<string> ctNameList = new HashSet<string>();

            foreach (PlanSetup planSetup in course.PlanSetups) {
                if (planSetup.IsTreated == true) {
                    if (planSetup.StructureSet == null) {
                        throw new ArgumentNullException($"StructureSet in {planSetup.Id} is null");
                    }

                    if(planSetup.NumberOfFractions == null) {
                        throw new ArgumentException("null in number of fraction");
                    }


                    string ctName = planSetup.StructureSet.Id;
                    ctNameList.Add(ctName);
                }

            }

            if (ctNameList.Count() != 1) {
                return false;
            }
            return true;

        }
        static void ExportFractionFromCSV(Application app) {

            using (var reader = new StreamReader(CSVPath)) {

                while (!reader.EndOfStream) {
                    var line = reader.ReadLine();
                    string[] columes = line.Split(',');
                    try {

                        ExportFraction(app, columes[0], columes[1], columes[2]);


                    } catch (Exception e) {
                        PatientIDList.Add(columes[0]);
                        Console.Error.WriteLine(e.ToString());
                        app.ClosePatient();
                    }
                    ExportErrorFile(PatientIDList, ExportFolder + "\\error.txt");
                    Console.WriteLine($"Finish ...");

                }
            }




        }

        private static void ExportErrorFile(List<string> list, string path) {

            using (TextWriter tw = new StreamWriter(path)) {
                foreach (string s in list) {
                    tw.WriteLine(s);
                }
            }
        }

        static void ExportFraction(Application app, string patientId, string courseId, string lungName) {
            Console.WriteLine($"Reading {patientId}");
            Patient patient = app.OpenPatientById(patientId);
            Course course = patient.Courses.Single(c => c.Id == courseId);

            if (!IsCTImageOK(course)) {
                throw new Exception("There are more than one Registered CT image");
            }



            List<TreatmentFraction> treatmentFractions = new List<TreatmentFraction>();
            TreatmentDict treatmentDict = AddFraction(treatmentFractions, course);

            StructureSet structureSet = treatmentFractions[0].PlanSetupList[0].StructureSet;
            Structure lungStruct = structureSet.Structures.Single(s => s.Id == lungName);
            Image ct = structureSet.Image;

            if (IsExport) {
                string exportFolder = ExportFolder + "\\" + patient.Id;
                CreateFolder(exportFolder);

                ExportImage(ct, exportFolder);
                ExportDose(treatmentFractions, treatmentDict, exportFolder);
                ExportNumberOfFraction(treatmentFractions, treatmentDict, exportFolder);
                ExportStructure(lungStruct, exportFolder, "\\lung.ply");
            }


            app.ClosePatient();

        }

        static void ExportStructure(Structure structure, string folder, string name) {
            if (!structure.HasSegment) {
                Console.WriteLine($"NO SEGMENTED");
                return;
            }
            Console.WriteLine($"HAVE SEGMENTED");
            Console.WriteLine($"Exporting Structure ...");
            //comment out the file output you do not want 
            SaveTriangleMeshToPlyFile(structure.MeshGeometry, folder + name);

        }

        static public void ExportNumberOfFraction(List<TreatmentFraction> treatmentFractions, TreatmentDict treatmentDict, string exportFolder) {

            string outputFileName = exportFolder + "//fraction.txt";
            string name;
            if (File.Exists(outputFileName)) {
                File.SetAttributes(outputFileName, FileAttributes.Normal);
                File.Delete(outputFileName);
            }
            using (TextWriter writer = new StreamWriter(outputFileName)) {
                foreach (TreatmentFraction treatmentFraction in treatmentFractions) {
                    writer.Write(treatmentFraction.NumberOfFraction.ToString() + " ");
                    writer.Write(treatmentFraction.PlanSetupList.Count().ToString() + " ");
                    foreach(PlanSetup planSetup in treatmentFraction.PlanSetupList) {
                        name = treatmentDict.PlanDict[planSetup.Id];
                        writer.Write(name + " ");
                    }
                    writer.WriteLine();

                }
            }


        }

        static public void ExportImage(Image image, string folder) {
            Console.WriteLine($"Exporting CT image ...");
            string filename = folder + "\\ct.vtk";

            SaveDoseOrImageToVTKStructurePoints(null, image, filename);
        }
        static public void ExportDose(List<TreatmentFraction> treatmentFractions, TreatmentDict treatmentDict, string exportFolder) {
            string name;
            foreach (TreatmentFraction treatmentFraction in treatmentFractions) {
                foreach (PlanSetup planSetup in treatmentFraction.PlanSetupList) {
                    Dose fractionDose = planSetup.Dose;
                    if (!planSetup.DoseValuePresentation.ToString().Contains("Absolute")) {

                        int maxValueForScaling = fractionDose != null ? FindMaxValue(fractionDose) : 0;

                        double? constant = (double)planSetup.TotalDose.Dose / 100 / planSetup.PlanNormalizationValue / planSetup.NumberOfFractions / planSetup.TreatmentPercentage;
                        double? test = maxValueForScaling * constant;
                        if (constant == null) {
                            throw new Exception("Constant is null");
                        }

                        name = treatmentDict.PlanDict[planSetup.Id];
                        ExportDoseFraction(fractionDose, exportFolder + $"//{name}.vtk", constant);
                    }

                    // --------------------------------------------------------------------------------------------- 
                    // I did not written the code for export absolute dose because our data are record in relative dose
                    //--------------------------------------------------------------------------------------------- 
                    name = treatmentDict.PlanDict[planSetup.Id];
                    ExportDoseFraction(fractionDose, exportFolder + $"//{name}.vtk", 1);
                }
            }

        }

        static public void ExportDoseFraction(Dose dose, string outputFileName, double? constant) {
            if (File.Exists(outputFileName)) {
                File.SetAttributes(outputFileName, FileAttributes.Normal);
                File.Delete(outputFileName);
            }

            int W, H, D;
            double sx, sy, sz;
            VVector origin, rowDirection, columnDirection;

            W = dose.XSize;
            H = dose.YSize;
            D = dose.ZSize;
            sx = dose.XRes;
            sy = dose.YRes;
            sz = dose.ZRes;
            origin = dose.Origin;
            rowDirection = dose.XDirection;
            columnDirection = dose.YDirection;

            using (TextWriter writer = new StreamWriter(outputFileName)) {
                writer.WriteLine("# vtk DataFile Version 3.0");
                writer.WriteLine("vtk output");
                writer.WriteLine("ASCII");
                writer.WriteLine("DATASET STRUCTURED_POINTS");
                writer.WriteLine("DIMENSIONS " + W + " " + H + " " + D);

                int[,] buffer = new int[W, H];

                double xsign = rowDirection.x > 0 ? 1.0 : -1.0;
                double ysign = columnDirection.y > 0 ? 1.0 : -1.0;
                double zsign = GetZDirection(rowDirection, columnDirection).z > 0 ? 1.0 : -1.0;

                writer.WriteLine("ORIGIN " + origin.x.ToString() + " " + origin.y.ToString() + " " + origin.z.ToString());
                writer.WriteLine("SPACING " + sx * xsign + " " + sy * ysign + " " + sz * zsign);
                writer.WriteLine("POINT_DATA " + W * H * D);
                writer.WriteLine("SCALARS image_data unsigned_short 1");
                writer.WriteLine("LOOKUP_TABLE default");
                //int maxValueForScaling = dose != null ? FindMaxValue(dose) : 0;

                for (int z = 0; z < D; z++) {
                    dose.GetVoxels(z, buffer);

                    for (int y = 0; y < H; y++) {
                        for (int x = 0; x < W; x++) {
                            int value = buffer[x, y];
                            UInt16 curvalue = 0;

                            curvalue = (UInt16)((double)(value / 10) * constant);
                            writer.Write(curvalue + " ");
                        }
                        writer.WriteLine();
                    }
                }
            }
        }
        static void SaveTriangleMeshToPlyFile(MeshGeometry3D mesh, string outputFileName) {
            if (mesh == null)
                return;

            if (File.Exists(outputFileName)) {
                File.SetAttributes(outputFileName, FileAttributes.Normal);
                File.Delete(outputFileName);
            }

            Point3DCollection vertexes = mesh.Positions;
            Int32Collection indexes = mesh.TriangleIndices;

            using (TextWriter writer = new StreamWriter(outputFileName)) {
                writer.WriteLine("ply");
                writer.WriteLine("format ascii 1.0");
                writer.WriteLine("element vertex " + vertexes.Count);

                writer.WriteLine("property float x");
                writer.WriteLine("property float y");
                writer.WriteLine("property float z");
                writer.WriteLine("property float nx");
                writer.WriteLine("property float ny");
                writer.WriteLine("property float nz");

                writer.WriteLine("element face " + indexes.Count / 3);

                writer.WriteLine("property list uchar int vertex_indices");

                writer.WriteLine("end_header");

                for (int v = 0; v < vertexes.Count(); v++) {
                    Vector3D normal = CalculateVertexNormal(mesh, v);

                    writer.Write(vertexes[v].X.ToString("e") + " ");
                    writer.Write(vertexes[v].Y.ToString("e") + " ");
                    writer.Write(vertexes[v].Z.ToString("e") + " ");
                    writer.Write(normal.X.ToString("e") + " ");
                    writer.Write(normal.Y.ToString("e") + " ");
                    writer.Write(normal.Z.ToString("e"));

                    writer.WriteLine();
                }

                int i = 0;
                while (i < indexes.Count) {
                    writer.Write("3 ");
                    writer.Write(indexes[i++] + " ");
                    writer.Write(indexes[i++] + " ");
                    writer.Write(indexes[i++] + " ");
                    writer.WriteLine();
                }
            }
        }
        static Vector3D CalculateSurfaceNormal(Point3D p1, Point3D p2, Point3D p3) {
            Vector3D v1 = new Vector3D(0, 0, 0); // Vector 1 (x,y,z) & Vector 2 (x,y,z) 
            Vector3D v2 = new Vector3D(0, 0, 0);
            Vector3D normal = new Vector3D(0, 0, 0);

            // Finds The Vector Between 2 Points By Subtracting 
            // The x,y,z Coordinates From One Point To Another. 

            // Calculate The Vector From Point 2 To Point 1 
            v1.X = p1.X - p2.X; // Vector 1.x=Vertex[0].x-Vertex[1].x 
            v1.Y = p1.Y - p2.Y; // Vector 1.y=Vertex[0].y-Vertex[1].y 
            v1.Z = p1.Z - p2.Z; // Vector 1.z=Vertex[0].y-Vertex[1].z 
                                // Calculate The Vector From Point 3 To Point 2 
            v2.X = p2.X - p3.X; // Vector 1.x=Vertex[0].x-Vertex[1].x 
            v2.Y = p2.Y - p3.Y; // Vector 1.y=Vertex[0].y-Vertex[1].y 
            v2.Z = p2.Z - p3.Z; // Vector 1.z=Vertex[0].y-Vertex[1].z 

            // Compute The Cross Product To Give Us A Surface Normal 
            normal.X = v1.Y * v2.Z - v1.Z * v2.Y; // Cross Product For Y - Z 
            normal.Y = v1.Z * v2.X - v1.X * v2.Z; // Cross Product For X - Z 
            normal.Z = v1.X * v2.Y - v1.Y * v2.X; // Cross Product For X - Y 

            normal.Normalize();

            return normal;
        }

        //creates a normal for a single vertex by searching all faces it is connected with 
        //then averaging the surface normal for those faces 
        static Vector3D CalculateVertexNormal(MeshGeometry3D mesh, int vertex) {
            Vector3D normal = new Vector3D(0, 0, 0);
            List<Vector3D> normals = new List<Vector3D>();

            //foreach triangle 
            for (int i = 0; i < mesh.TriangleIndices.Count(); i += 3) {
                //foreach vertex 
                for (int ii = 0; ii < 3; ii++) {
                    //calculates and add the surface normal if the face uses that vertex 
                    if (mesh.TriangleIndices[i + ii] == vertex) {
                        Vector3D surfaceNormal = CalculateSurfaceNormal(mesh.Positions[mesh.TriangleIndices[i]], mesh.Positions[mesh.TriangleIndices[i + 1]], mesh.Positions[mesh.TriangleIndices[i + 2]]);
                        normals.Add(surfaceNormal);
                    }
                }
            }

            //average the normals and normalize 
            foreach (Vector3D v in normals) {
                normal += v;
            }

            normal = normal / normals.Count();

            normal.Normalize();

            return normal;
        }


        static public void SaveDoseOrImageToVTKStructurePoints(Dose dose, Image image, string outputFileName) {
            if (File.Exists(outputFileName)) {
                File.SetAttributes(outputFileName, FileAttributes.Normal);
                File.Delete(outputFileName);
            }

            int W, H, D;
            double sx, sy, sz;
            VVector origin, rowDirection, columnDirection;
            if (dose != null) {
                W = dose.XSize;
                H = dose.YSize;
                D = dose.ZSize;
                sx = dose.XRes;
                sy = dose.YRes;
                sz = dose.ZRes;
                origin = dose.Origin;
                rowDirection = dose.XDirection;
                columnDirection = dose.YDirection;
            } else {
                W = image.XSize;
                H = image.YSize;
                D = image.ZSize;
                sx = image.XRes;
                sy = image.YRes;
                sz = image.ZRes;
                origin = image.Origin;
                rowDirection = image.XDirection;
                columnDirection = image.YDirection;
            }

            using (TextWriter writer = new StreamWriter(outputFileName)) {
                writer.WriteLine("# vtk DataFile Version 3.0");
                writer.WriteLine("vtk output");
                writer.WriteLine("ASCII");
                writer.WriteLine("DATASET STRUCTURED_POINTS");
                writer.WriteLine("DIMENSIONS " + W + " " + H + " " + D);

                int[,] buffer = new int[W, H];

                double xsign = rowDirection.x > 0 ? 1.0 : -1.0;
                double ysign = columnDirection.y > 0 ? 1.0 : -1.0;
                double zsign = GetZDirection(rowDirection, columnDirection).z > 0 ? 1.0 : -1.0;

                writer.WriteLine("ORIGIN " + origin.x.ToString() + " " + origin.y.ToString() + " " + origin.z.ToString());
                writer.WriteLine("SPACING " + sx * xsign + " " + sy * ysign + " " + sz * zsign);
                writer.WriteLine("POINT_DATA " + W * H * D);
                writer.WriteLine("SCALARS image_data unsigned_short 1");
                writer.WriteLine("LOOKUP_TABLE default");
                // int maxValueForScaling = dose != null ? FindMaxValue(dose) : 0;

                for (int z = 0; z < D; z++) {
                    if (dose != null) dose.GetVoxels(z, buffer);
                    else image.GetVoxels(z, buffer);
                    for (int y = 0; y < H; y++) {
                        for (int x = 0; x < W; x++) {
                            int value = buffer[x, y];
                            UInt16 curvalue = 0;
                            if (image != null)
                                curvalue = (UInt16)value;
                            else
                                curvalue = (UInt16)((double)(value / 10000));
                            writer.Write(curvalue + " ");
                        }
                        writer.WriteLine();
                    }
                }
            }
        }
        static int FindMaxValue(Dose dose) {
            int maxValue = 0;
            //float meanValue = 0;
            int[,] buffer = new int[dose.XSize, dose.YSize];
            if (dose != null) {
                for (int z = 0; z < dose.ZSize; z++) {
                    dose.GetVoxels(z, buffer);
                    for (int y = 0; y < dose.YSize; y++) {
                        for (int x = 0; x < dose.XSize; x++) {
                            int value = buffer[x, y];
                            //meanValue = meanValue + buffer[x, y];
                            if (value > maxValue)
                                maxValue = value;
                        }
                    }
                }
            }
            //meanValue = meanValue / (dose.ZSize * dose.XSize * dose.YSize);
            //return (int)meanValue;
            return maxValue;
        }
        private static void CreateFolder(string path) {

            try {
                if (Directory.Exists(path)) {
                    return;
                }
                DirectoryInfo di = Directory.CreateDirectory(path);
                Console.WriteLine($"CREATE {path}");
            } catch (Exception e) {
                Console.WriteLine($"process failed {e.ToString()}");

            }
        }
        private static VVector GetZDirection(VVector a, VVector b) {
            // return cross product 
            return new VVector(a.y * b.z - a.z * b.y, a.z * b.x - a.x * b.z, a.x * b.y - a.y * b.x);
        }
    }
}
