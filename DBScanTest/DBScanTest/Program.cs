using Spire.Xls;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DBScan
{
    class Program
    {
        static void Main(string[] args)
        {
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"C:\Users\Damla\source\repos\DBScanTest\DBScanTest\bin\Debug\Deneme.xlsx"); // Veri konumu
            Worksheet sheet = workbook.Worksheets[0];
            var dt = sheet.ExportDataTable();

            double epsilon = 1.5;
            int minPts = 3;

            // Butun data indexlendi
            int index = 0;
            Indexing(dt, index);

            // Labeling asamasi icin her dataya basta sifir atandi (belirsiz olduklari icin)
            dt.Columns.Add("Label");
            LabelingZero(dt);

            // Epsilon degeri tutan datalar icin ayri bir tablo duzenlendi
            DataTable dt2 = new DataTable();
            dt2.Columns.Add("Distance Name");
            dt2.Columns.Add("Value");

            // Fonksiyonlar
            EuclidCalculation(dt); // her nokta arasi oklid uzakligi hesaplandi
            EpsilonChecker(dt, epsilon); // epsilon degerine uyanlar belirlendi
            PrintEpsilonOK(dt, dt2); // epsilon değerine uyan uzaklıklar ayrı tabloda(dt2) gösterildi

            // Komsulari gostermek icin tablo yaratildi
            DataTable Neighbors = new DataTable();
            Neighbors.Columns.Add("Points");
            Neighbors.Columns.Add("Neighbors");
            InitializeNeighborTable(Neighbors, dt); 

            // Kumeleri birlestirmek icin ortak komsularin belirlendigi tablo yaratildi
            DataTable dt3 = new DataTable(); 
            dt3.Columns.Add("Points");
            dt3.Columns.Add("Neighbors");
            InitializeDt3(dt3, dt); // Pointler tabloya yerlestirildi

            for (int i = 0; i < dt.Rows.Count; i++) // labeli sifir olan pointleri gezmesi gerekiyor
            {
                var point = dt.Rows[i]["Points"].ToString();
                if (dt.Rows[i]["Label"].ToString().Equals("0")) 
                {
                    RegionFinder(Neighbors, dt, point, minPts);
                }
            }

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (int.Parse(dt.Rows[i]["Label"].ToString()) > 0) 
                {
                    string point = dt.Rows[i]["Points"].ToString(); 
                    //Console.WriteLine("noktamiz: " + point);
                    string pointNeighbor = "";
                    string pointLabel = dt.Rows[i]["Label"].ToString();
                    char[] neighborChar = Neighbors.Rows[i]["Neighbors"].ToString().ToCharArray(); 
                    //Console.WriteLine("komsulari: ");
                    for (int j = 0; j < neighborChar.Length; j++) 
                    {
                        pointNeighbor = neighborChar[j].ToString();
                       // Console.WriteLine(neighborChar[j]);
                        GrowCluster(dt, Neighbors, dt3, pointNeighbor, point, pointLabel, minPts); 
                    }
                }
            }

            string finalNeighbors = ""; 
            for (int j = dt3.Rows.Count-1 ; j >= 0; j--) 
            {
                foreach (DataRow dr in dt.Rows) //dt tablosunu gezip yerlestirecegiz
                {
                    if (dr["Points"].ToString().Equals(dt3.Rows[j]["Points"])) 
                    { 
                        finalNeighbors = dt3.Rows[j]["Neighbors"].ToString(); 
                        char[] final = finalNeighbors.ToCharArray();

                        for (int i = 0; i < final.Length; i++)
                        {
                            foreach (DataRow sdr in dt.Rows)
                            {
                                foreach (DataColumn sdc in dt.Columns)
                                {
                                    if (sdr["Points"].Equals(final[i].ToString()))
                                    {
                                    sdr["Label"] = dt.Rows[j]["Label"].ToString();
                                    }
                                }
                            }
                        }
                    }
                }
            }

            // Print test  [Buraya tablolari goruntuleyebilmek icin debug isareti koyun.]
            Console.WriteLine("Debug modunda(111. satir) dt tablosunun label kolonuna bakiniz.");
            Console.ReadKey();
        }

        /*
         *Main fonksiyon sonu 
         */

        // Dt tablosundaki veriler indexlendi
        public static void Indexing(DataTable dt, int index)
        {
            dt.Columns.Add("Index");
            foreach (DataRow item in dt.Rows)
            {
                foreach (DataColumn item2 in dt.Columns)
                {
                    item["Index"] = index;
                }
                index++;
            }
        }

        // Dt tablosundaki butun labellar sifira esitlendi
        public static void LabelingZero(DataTable dt)
        {
            int notDecided = 0; //Belirsiz olduklari icin sifir
            foreach (DataRow item in dt.Rows)
            {
                foreach (DataColumn item2 in dt.Columns)
                {
                    item["Label"] = notDecided;
                }
            }
        }

        // Butun verilerin birbirleriyle olan oklid uzakliklari hesaplandi
        public static void EuclidCalculation(DataTable dt)
        {
            string columnName;
            foreach (DataRow item in dt.Rows)
            {
                columnName = "distance " + item[0].ToString(); // kolon isimlendirmesi distance + "Point"
                dt.Columns.Add(columnName);

                foreach (DataRow item2 in dt.Rows)
                {
                    if (item.ItemArray != item2.ItemArray) // Kendisiyle karsilastirmayi engellemek icin
                    {
                        item2[columnName] = Math.Sqrt((Math.Pow((GetDoubleFromString(item["X1"]) - GetDoubleFromString(item2["X1"])), 2) +
                                                       Math.Pow((GetDoubleFromString(item["X2"]) - GetDoubleFromString(item2["X2"])), 2)));
                    }
                }
            }
        }

        // Tablodaki degerleri oklid hesaplamasina sokabilmek icin String'den Double degere cevirtiyoruz.
        public static double GetDoubleFromString(object value)
        {
            var str = value.ToString().Trim().Replace(".", ",");

            if (str.ToString().Split(',').Length == 2)
            {
                double d = Convert.ToDouble(str, CultureInfo.GetCultureInfo("tr-TR"));
                return d;
            }
            else if (str.ToString().Split(',').Length == 3)
            {
                var arr = str.Split(',');
                var newStr = arr[0] + arr[1] + ',' + arr[2];
                double d = Convert.ToDouble(newStr, CultureInfo.GetCultureInfo("tr-TR"));
                return d;
            }
            else
            {
                double d = Convert.ToDouble(str, CultureInfo.GetCultureInfo("tr-TR"));
                return d;
            }
        }

        // epsilon degerine esit veya kucuk mu kontrol ediyoruz.
        public static void EpsilonChecker(DataTable dt, double epsilon)
        {
            String seperator = "$|OK|$"; // Epsilon degeri icerisinde bulunanlari ayirt edebilmek icin basina seperator ekliyoruz
            foreach (DataRow dr in dt.Rows)
            {
                foreach (DataColumn dc in dt.Columns)
                {
                    if (dc.ColumnName.StartsWith("distance")) 
                    {
                        var value = dr[dc];
                        if (GetDoubleFromString(dr[dc].ToString()) <= epsilon && GetDoubleFromString(dr[dc].ToString()) != 0)
                        {
                            dr[dc] = seperator + dr[dc];
                        }
                    }
                }
            }
        }

        // epsilon degerine esit veya kucuk olanlari tabloya ekliyoruz
        public static void PrintEpsilonOK(DataTable dt, DataTable dt2)
        {
            foreach (DataRow dr in dt.Rows)
            {
                foreach (DataColumn dc in dt.Columns)
                {
                    var cellValue = dr[dc].ToString();
                    if (cellValue.StartsWith("$|OK|$"))
                    {
                        dt2.Rows.Add((new Object[] { "distance(" + dr[0] + "," + dc.ColumnName.Replace("distance ", "") + ")",
                                                     cellValue.Replace("$|OK|$","") }));
                    }
                }
            }
        }

        // tabloya pointleri olmasi gerektigi kadar setliyoruz (Asil dt tablosunda ne kadarsa o kadar point initiliaze ettik)
        public static void InitializeNeighborTable(DataTable Neighbors, DataTable dt)
        {
            foreach (DataRow dr in dt.Rows)
            {
                foreach (DataColumn dc in dt.Columns)
                {
                    if (dc.ColumnName.Equals("Points"))
                    {
                        Neighbors.Rows.Add(dr[dc].ToString());
                    }
                }
            }
        }

        public static void InitializeDt3(DataTable dt3, DataTable dt)
        {
            foreach (DataRow dr in dt.Rows)
            {
                foreach (DataColumn dc in dt.Columns)
                {
                    if (dc.ColumnName.Equals("Points"))
                    {
                        dt3.Rows.Add(dr[dc].ToString());
                    }
                }
            }
        }

        static int currentCluster = 0;
        // komsulari buluyoruz
        public static void RegionFinder(DataTable Neighbors, DataTable dt, string point, int minPts)
        {
            int counter = 1;
            string neighborStr = "";
            int indexing = 0;
            foreach (DataRow dr in dt.Rows)
            {
                if (!dr["Label"].Equals("0")) continue; // Ziyaret edilmemişlere bakıyoruz 
                if (dr["Points"].Equals(point)) // istenilen noktayi incelemek icin; mesela A
                {
                    foreach (DataColumn dc in dt.Columns) // kolonlari gezmeye basliyoruz
                    {
                        var cellValue = dr[dc].ToString(); // Current degere bakiyoruz
                        indexing = int.Parse(dr["Index"].ToString());
                        if (cellValue.StartsWith("$|OK|$"))
                        {
                            counter++; //komsu sayisi
                            neighborStr = neighborStr + dc.ColumnName.ToString().Replace("distance ", ""); // seperator ile baslayanlari komsu olarak belirledik
                            Neighbors.Rows[indexing]["Neighbors"] = neighborStr; // komsulari pointlerin yanina yerlestirdik
                        }
                    }
                    if (counter < minPts) // noise olanlar (kesin olmamakla beraber)
                        dr["Label"] = -1;
                    else
                    {
                        currentCluster++;
                        dr["Label"] = currentCluster; // buraya kadar core pointler belirlendi, -1 olup border kalanlar secilecek
                    }
                }
            }
        }

        public static void GrowCluster(DataTable dt, DataTable Neighbors, DataTable dt3, string pointNeighbor, string point, string pointLabel, int minPts) 
        {                                                                                                    
            foreach (DataRow dr in dt.Rows)
            {
                if (dr["Points"].Equals(pointNeighbor))
                {
                    if (int.Parse(dr["Label"].ToString()) == (-1) || int.Parse(dr["Label"].ToString()) == 0) //sifir veya -1'se labella
                    {
                        foreach (DataColumn dc in dt.Columns)
                        {
                            dr["Label"] = pointLabel;
                        }
                    }
                    else if (int.Parse(dr["Label"].ToString()) > 0)
                    {
                        // bu durumda komsu noktayi pointle birlestir                    
                        dr["Label"] = pointLabel;
                    }
                }
            }
            PnRegionFinder(dt, Neighbors, dt3, point, pointNeighbor, pointLabel);
        }

        public static void PnRegionFinder(DataTable dt, DataTable Neighbors, DataTable dt3, string point, string pointNeighbor, string pointLabel)
        {
            int index = dt.Rows.Count;
            foreach (DataRow dr in Neighbors.Rows)
            {
                if (dr["Points"].ToString().Equals(pointNeighbor)) // Komsu tablosunda komsu noktamizi bulduk
                {
                    // Console.WriteLine("Noktamizin kendisinin labeli: " + pointLabel);
                    char[] neighbors = dr["Neighbors"].ToString().ToCharArray();

                    for(int i= 0; i<neighbors.Length; i++) 
                    { 
                        foreach (DataRow dr3 in dt3.Rows)
                        {
                            foreach (DataColumn dc3 in dt3.Columns)
                            {
                                if (dr3["Points"].Equals(point))
                                {
                                    if (dr3["Neighbors"].ToString().ToCharArray().Contains(neighbors[i]))
                                        continue;
                                    dr3["Neighbors"] = dr3["Neighbors"].ToString() + neighbors[i];
                                }
                            }
                        }
                    }
                }
            }
        }

    }
}


/* Functions:
* Indexing
* Labeling Zero
* Euclid Calculation
* Epsilon Checker
* Print Epsilon OK
* initializeNeighborTable
* Get Double from String
*/
