using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using Microsoft.Office.Interop.Excel;
using ClosedXML.Excel;
using System.Data.SqlClient;
using System.IO;
using System.Net.Mail;
using System.Reflection;

namespace DSR
{
    class Program
    {
        private static object oMissing;

        static void Main(string[] args)
        {
            //Console.WriteLine("11111111111111111111111" + DateTime.Now);
            //Console.ReadKey();

            string connectionString = @"Data Source=192.168.0.5;Initial Catalog=FeroApp;User ID=sa;Password=AdminFX9.";


            string c7 = "0", c8 = "0", c9 = "0", c10 = "0", c11 = "0", c12 = "0", c13 = "0", c14 = "0", c15 = "0", c16 = "0", c17 = "0" , c18 = "0" , c19 = "0", c21 = "0"; // 
            double f7 = 0.0, f8 = 0.0, f9 = 0.0, f10 = 0.0, f11 = 0.0, f12 = 0.0, f13 = 0.0, f14 = 0.0, f15 = 0.0, f16 = 0.0, f17 = 0.0,  f18=0.0 , f19 = 0.0, f21 = 0.0; // 
            double g7 = 0.0, g8 = 0.0, g9 = 0.0, g10 = 0.0, g11 = 0.0, g12 = 0.0, g13 = 0.0, g14 = 0.0, g15 = 0.0, g16 = 0.0, g17 = 0.0,  g18=0.0 , g19 = 0.0, g21 = 0.0; // 

            double i7 = 0.0, i8 = 0.0, i9 = 0.0, i10 = 0.0, i11 = 0.0, i12 = 0.0, i13 = 0.0, i14 = 0.0, i15 = 0.0, i16 = 0.0, i17 = 0   , i18=0   , i19 = 0.0, i21 = 0.0 , i22=0.0; // 
            string j7 = "0", j8 = "0", j9 = "0", j10 = "0", j11 = "0", j12 = "0", j13 = "0", j14 = "0", j15 = "0", j16 = "0", j17="0"   , j18 ="0",j19 = "0", j21 = "0"; // 
            double k7 = 0.0, k8 = 0.0, k9 = 0.0, k10 = 0.0, k11 = 0.0, k12 = 0.0, k13 = 0.0, k14 = 0.0, k15 = 0.0, k16 = 0.0, k17=0.0   , k18 = 0, k19 = 0.0, k21 = 0; // 
            double q7 = 0.0, q8 = 0.0, q9 = 0.0, q10 = 0.0, q11 = 0.0, q12 = 0.0, q13 = 0.0, q14 = 0.0, q15 = 0.0, q16 = 0.0, q17=0.0   , q18 = 0, q19 = 0.0, q21 = 0; // 
            string sql1 = "";
            string dat1 = "2017-06-20", dat2 = "2017-06-20", dat3 = "";
            string vo = "";

            DateTime d1 = new DateTime(2017, 6, 20);
            DateTime d10 = new DateTime(2017, 7, 27);
            DateTime d2 = DateTime.Now.AddDays(-1);
            DateTime d20 = d2;
            //DateTime d2 = DateTime.Now;
            if (d2.Hour < 12)
            {
//                           d2 = d2.AddDays(-1);
            }
            DateTime d3 = DateTime.Now;

            if (1 == 2)
            {
                d2 = new DateTime(2017, 9, 16);
                d3 = new DateTime(2017, 9, 17);
            }

            string dan1 = d2.Day.ToString();
            string m1 = d2.Month.ToString();
            string g1 = d2.Year.ToString();

            dat1 = d2.Year.ToString() + '-' + d2.Month.ToString() + '-' + d2.Day.ToString();
            dat2 = dat1;
            string datLDP = dat1;

            dat3 = d3.Year.ToString() + '-' + d3.Month.ToString() + '-' + d3.Day.ToString();

            TimeSpan t = d2 - d1;
            int dana = t.Days;
            string nuland = "", nulanm = "", dnuland = "", dnulanm = "";
            if (d2.Day <= 9)
            {
                nuland = "0";
            }
            if (d2.Month <= 9)
            {
                nulanm = "0";
            }

            if (d3.Day <= 9)
            {
                dnuland = "0";
            }
            if (d3.Month <= 9)
            {
                dnulanm = "0";
            }

            string dat10 = nuland + d2.Day.ToString() + '.' + nulanm + d2.Month.ToString() + '.' + d2.Year.ToString();
            string dat100 = dat10;
            string datLDPTitle = dat10;
            string dat30 = dnuland + d3.Day.ToString() + '.' + dnulanm + d3.Month.ToString() + '.' + d3.Year.ToString();   // današnji datum

            //string fileName = @"L:\izvještaji\dsr\dsr"+ nuland+d2.Day.ToString() + nulanm + d2.Month.ToString() + d2.Year.ToString() + ".xlsm";
            string fileName = @"c:\brisi\dsr" + nuland + d2.Day.ToString() + nulanm + d2.Month.ToString() + d2.Year.ToString() + ".xlsx";
            Console.WriteLine("Od datuma: " + dat1 + " trenutno vrijeme " + DateTime.Now);
            double ukupv = 0.0, ukupk = 0.0;
            DateTime input = d2;
            int delta = DayOfWeek.Monday - input.DayOfWeek;
            if (d2.DayOfWeek == DayOfWeek.Sunday)
            {
                delta = -6;
            }
            DateTime monday = input.AddDays(delta);
            delta = DayOfWeek.Sunday - input.DayOfWeek + 7;
            if (d2.DayOfWeek == DayOfWeek.Sunday)
            {
                delta = 0;
            }

            string datreps = dat10;   // jučerašnji datum
            DateTime sunday = input.AddDays(delta);
            //monday= new DateTime(2017, 7, 31);
            sunday = monday.AddDays(6);
            sunday = d2;

            //Application excel = new Application();
            //Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            //Workbook workbook = excel.Workbooks.Open(@"C:\brisi\dsr_template33.xlsx", ReadOnly: false, Editable: true);
            //Workbook workbook = excel.Workbooks.Open(@"C:\brisi\dsr_template1001.xlsx", ReadOnly: false, Editable: true);
            //Workbook workbook = excel.Workbooks.Open(@"C:\brisi\dsr_template2302.xlsx", ReadOnly: false, Editable: true);
            var workbook = new XLWorkbook(@"C:\brisi\dsr_template2103.xlsx");

            /*On the worksheet, create the worksheet in another sheet named user reports,   
             * this leaflet will be included in the excel file which will then be generated.*/
//            var worksheet = workbook.Worksheets.Add(@"C:\brisi\dsr_template2103.xlsx"); //, ReadOnly: false, Editable: true);


            for (int ws = 1; ws <= 4; ws++)
            {

                if (ws == 2)
                {
                    dat10 = monday.Day.ToString() + "." + monday.Month.ToString() + "." + monday.Year.ToString() + " - " + sunday.Day.ToString() + "." + sunday.Month.ToString() + "." + sunday.Year.ToString();
                    dat1 = monday.Year.ToString() + "-" + monday.Month.ToString() + "-" + monday.Day.ToString();
                    dat2 = sunday.Year.ToString() + "-" + sunday.Month.ToString() + "-" + sunday.Day.ToString();
                }
                else
                {
                    dat10 = dat100;
                    d2 = d20;
                    dat1 = d2.Year.ToString() + '-' + d2.Month.ToString() + '-' + d2.Day.ToString();
                    dat2 = dat1;                    
                }

                if (1 == 1)
                {
                    // ostvareno komada
                    using (SqlConnection cn = new SqlConnection(connectionString))
                    {
                        cn.Open();
                        //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                        // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);

                        sql1 = "rfind.dbo.realizacija '" + dat1 + "','" + dat2 + "',1";

                        SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                        SqlDataReader reader = sqlCommand.ExecuteReader();
                        string kupac;

                        while (reader.Read())
                        {
                            kupac = reader["Kupac"].ToString();
                            if (kupac.Contains("Austria"))
                            {
                                c7 = (reader["Kolicinaok"].ToString());
                                //j7 = (reader["broj_stelanja"].ToString());  // BDF
                                ukupv = ukupv + (double.Parse)(reader["vrijednost"].ToString());
                                ukupk = ukupk + (double.Parse)(reader["kolicinaok"].ToString());
                            }

                            if (kupac.Contains("SCHWEINFURT"))
                            {
                                c8 = (reader["Kolicinaok"].ToString());
                                //j8 = (reader["broj_stelanja"].ToString());  // BMW
                                ukupv = ukupv + (double.Parse)(reader["vrijednost"].ToString());
                                ukupk = ukupk + (double.Parse)(reader["kolicinaok"].ToString());
                            }

                            if (kupac.Contains("FAG"))
                            {
                                c9 = (reader["Kolicinaok"].ToString());
                                //j9 = (reader["broj_stelanja"].ToString());  // DEbrecin
                                ukupv = ukupv + (double.Parse)(reader["vrijednost"].ToString());
                                ukupk = ukupk + (double.Parse)(reader["kolicinaok"].ToString());
                            }

                            if (kupac.Contains("ROMANIA"))
                            {
                                c10 = (reader["Kolicinaok"].ToString());
                                //j10 = (reader["broj_stelanja"].ToString());  // Brašov
                                ukupv = ukupv + (double.Parse)(reader["vrijednost"].ToString());
                                ukupk = ukupk + (double.Parse)(reader["kolicinaok"].ToString());
                            }

                            if (kupac.Contains("Technologies"))  // wupertal
                            {
                                if (kupac.Contains("Prsteni"))  // wupertal
                                {
                                    c11 = (reader["Kolicinaok"].ToString());
                                    //  j11 = (reader["broj_stelanja"].ToString());  // 
                                    ukupv = ukupv + (double.Parse)(reader["vrijednost"].ToString());
                                    ukupk = ukupk + (double.Parse)(reader["kolicinaok"].ToString());
                                }
                            }

                            if (kupac.Contains("Technologies"))  // wupertal
                            {

                                if (kupac.Contains("Valjci"))  // wupertal
                                {
                                    c12 = (reader["Kolicinaok"].ToString());
                                    //j12 = (reader["broj_stelanja"].ToString());  // 
                                    ukupv = ukupv + (double.Parse)(reader["vrijednost"].ToString());
                                    ukupk = ukupk + (double.Parse)(reader["kolicinaok"].ToString());
                                }
                            }

                            if (kupac.Contains("KYSUCE"))  // 
                            {
                                if (kupac.Contains("Ku"))  // 
                                {
                                    c13 = (reader["Kolicinaok"].ToString());
                                    //j13 = (reader["broj_stelanja"].ToString());  // 
                                    ukupv = ukupv + (double.Parse)(reader["vrijednost"].ToString());
                                    ukupk = ukupk + (double.Parse)(reader["kolicinaok"].ToString());
                                }
                            }

                            if (kupac.Contains("KYSUCE"))  // 
                            {
                                if (kupac.Contains("Prsten"))  // 
                                {
                                    c17 = (reader["Kolicinaok"].ToString());
                                    //j13 = (reader["broj_stelanja"].ToString());  // 
                                    ukupv = ukupv + (double.Parse)(reader["vrijednost"].ToString());
                                    ukupk = ukupk + (double.Parse)(reader["kolicinaok"].ToString());
                                }
                            }


                            if (kupac.Contains("FERRO"))  // wupertal
                            {
                                c14 = (reader["Kolicinaok"].ToString());
                                //j14 = (reader["broj_stelanja"].ToString());  // 
                                ukupv = ukupv + (double.Parse)(reader["vrijednost"].ToString());
                                ukupk = ukupk + (double.Parse)(reader["kolicinaok"].ToString());
                            }

                            if (kupac.Contains("SONA"))
                            {
                                c15 = (reader["Kolicinaok"].ToString());
                                //j15 = (reader["broj_stelanja"].ToString());  // 
                                ukupv = ukupv + (double.Parse)(reader["vrijednost"].ToString());
                                ukupk = ukupk + (double.Parse)(reader["kolicinaok"].ToString());
                            }

                            if (kupac.Contains("SIGMA"))  // wupertal
                            {
                                c16 = (reader["Kolicinaok"].ToString());
                                //j16 = (reader["broj_stelanja"].ToString());  // 
                                ukupv = ukupv + (double.Parse)(reader["vrijednost"].ToString());
                                ukupk = ukupk + (double.Parse)(reader["kolicinaok"].ToString());
                            }

                            if (kupac.Contains("TECHNOLOGIES"))  // ELTMANN
                            {
                                c18 = (reader["Kolicinaok"].ToString());
                                //j16 = (reader["broj_stelanja"].ToString());  // 
                                ukupv = ukupv + (double.Parse)(reader["vrijednost"].ToString());
                                ukupk = ukupk + (double.Parse)(reader["kolicinaok"].ToString());
                            }

                            if (kupac.Contains("Brasil"))  // Brasil
                            {
                                c19 = (reader["Kolicinaok"].ToString());
                                //j19 = (reader["broj_stelanja"].ToString());  // 
                                ukupv = ukupv + (double.Parse)(reader["vrijednost"].ToString());
                                ukupk = ukupk + (double.Parse)(reader["kolicinaok"].ToString());
                            }

                            if (kupac.Contains("NEUWEG"))  // NSK
                            {
                                c21 = (reader["Kolicinaok"].ToString());
                                //j19 = (reader["broj_stelanja"].ToString());  // 
                                ukupv = ukupv + (double.Parse)(reader["vrijednost"].ToString());
                                ukupk = ukupk + (double.Parse)(reader["kolicinaok"].ToString());
                            }

                        }
                        cn.Close();

                    }

                    Console.WriteLine("Izračunate količine   trenutno vrijeme " + DateTime.Now);
                } // end if 1=2

                // broj štelanja po novome
                j7 = "0"; j8 = "0"; j9 = "0"; j10 = "0"; j11 = "0"; j12 = "0"; j13 = "0"; j14 = "0"; j15 = "0"; j16 = "0"; j17 = "0" ; j19 = "0"; j21 = "0"; // 
                using (SqlConnection cn = new SqlConnection(connectionString))
                {
                    cn.Open();
                    //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id;[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                    // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                    if (ws <= 2)
                    {
                        sql1 = "rfind.dbo.LDP_recalc '" + dat1 + "','" + dat3 + "',102";
                    }
                    else
                    {
                        sql1 = "rfind.dbo.LDP_recalc '" + dat1 + "','" + dat3 + "',1020";
                    }

                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string kupac, vrstanarudzbe,hala1="";

                    while (reader.Read())
                    {
                        if (ws >= 3)
                        {
                            hala1 = reader["hala"].ToString();
                            if ((ws == 3) && hala1.Contains("3")) // pušta samo halu 1
                            {
                                continue;
                            }
                            if ((ws == 4) && hala1.Contains("1"))  // pušta samo 3
                            {
                                continue;
                            }
                        }


                        kupac = reader["Kupac"].ToString();
                        vrstanarudzbe = reader["vrstanarudzbe"].ToString();
                        if (kupac.Contains("Austria"))
                        {

                            j7 = (reader["broj_stelanja"].ToString());  // BDF

                        }

                        if (kupac.Contains("SCHWEINFURT"))
                        {

                            j8 = (reader["broj_stelanja"].ToString());  // BMW

                        }

                        if (kupac.Contains("FAG"))
                        {

                            j9 = (reader["broj_stelanja"].ToString());  // DEbrecin

                        }

                        if (kupac.Contains("ROMANIA"))
                        {

                            j10 = (reader["broj_stelanja"].ToString());  // Brašov

                        }

                        if (kupac.Contains("Technologies"))  // wupertal
                        {
                            if (vrstanarudzbe.Contains("Prsteni"))  // wupertal
                            {

                                j11 = (reader["broj_stelanja"].ToString());  // 

                            }
                        }

                        if (kupac.Contains("Technologies"))  // wupertal
                        {

                            if (vrstanarudzbe.Contains("Valjci"))  // wupertal
                            {

                                j12 = (reader["broj_stelanja"].ToString());  // 

                            }
                        }

                        if (kupac.Contains("KYSUCE"))  // 
                        {
                            if (vrstanarudzbe.Contains("Ku"))  // w
                            {
                                j13 = (reader["broj_stelanja"].ToString());  // 
                            }
                        }
                        if (kupac.Contains("KYSUCE"))  // 
                        {
                            if (vrstanarudzbe.Contains("Prsten"))  // w
                            {
                                j17 = (reader["broj_stelanja"].ToString());  // 
                            }
                        }


                        if (kupac.Contains("FERRO"))  // wupertal
                        {

                            j14 = (reader["broj_stelanja"].ToString());  // 

                        }

                        if (kupac.Contains("SONA"))
                        {

                            j15 = (reader["broj_stelanja"].ToString());  // 

                        }

                        if (kupac.Contains("SIGMA"))  // wupertal
                        {

                            j16 = (reader["broj_stelanja"].ToString());  // 

                        }

                        if (kupac.Contains("TECHNOLOG"))  // eltman
                        {

                            j18 = (reader["broj_stelanja"].ToString());  // 

                        }

                        if (kupac.Contains("Brasil"))  // Brasil
                        {

                            j19 = (reader["broj_stelanja"].ToString());  // 

                        }
                        if (kupac.Contains("NEUWEG"))
                        {

                            j21 = (reader["broj_stelanja"].ToString());  // BDF

                        }


                    }
                    cn.Close();

                }
                Console.WriteLine("Izračunat broj štelanja  trenutno vrijeme " + DateTime.Now);


                // kupac, komada iz sELECT * FROM EvidencijaProizvedenoFakturirano_Zbirno('2017-08-02', '2017-08-02')
                c7 = ""; c8 = ""; c9 = ""; c10 = ""; c11 = ""; c12 = ""; c13 = ""; c14 = ""; c16 = ""; c17 = "";c18 = ""; c19 = ""; c21 = ""; // 
                f7 = 0.0; f8 = 0.0; f9 = 0.0; f10 = 0.0; f11 = 0.0; f12 = 0.0; f13 = 0.0; f14 = 0.0; f15 = 0.0; f16 = 0.0; f17 = 0.0; f18 = 0.0;f19 = 0.0; f21 = 0.0; // 
                g7 = 0.0; g8 = 0.0; g9 = 0.0; g10 = 0.0; g11 = 0.0; g12 = 0.0; g13 = 0.0; g14 = 0.0; g16 = 0.0; g17 = 0.0; g18 = 0.0; g19 = 0.0; // 

                string e7 = "", e8 = "", e9 = "", e10 = "", e11 = "", e12 = "", e13 = "", e14 = "", e15 = "", e16 = "", e17 = "", e18 = "", e19 = "", e21 = ""; // 

                ukupk = int.Parse(c15);    // količina od sone je ok, nju zadržavam
                ukupk = 0;

                using (SqlConnection cn = new SqlConnection(connectionString))
                {
                    cn.Open();
                    //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as iddoubl;[ime x] as ime;[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                    // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);

                    // SELECT * FROM feroapp.dbo.EvidencijaProizvedenoFakturirano_Zbirno(@dat1, @dat2)
                    if (ws <=2)
                    {
                        sql1 = "rfind.dbo.ldp_recalc '" + dat1 + "','" + dat2 + "',112";
                    }
                    else
                    {
                        sql1 = "rfind.dbo.ldp_recalc '" + dat1 + "','" + dat2 + "',113";
                    }

                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string kupac, vrstap,hala1="";


                    while (reader.Read())
                    {

                        if (ws >= 3)
                        {
                            hala1 = reader["hala"].ToString();
                            if ((ws == 3) && hala1.Contains("3")) // pušta samo halu 1
                            {
                                continue;
                            }
                            if ((ws == 4) && hala1.Contains("1"))  // pušta samo 3
                            {
                                continue;
                            }
                        }
                        kupac = reader["Kupac"].ToString();
                        vrstap = reader["vrstapro"].ToString();
                        vo = reader["vo"].ToString();


                        if (kupac.Contains("Austria"))
                        {
                            if (vo == "TOK")
                            {
                                c7 = (reader["kolicina"].ToString());
                                f7 = (double.Parse)(reader["vrijednost"].ToString());
                            }
                            else
                            {
                                e7 = (reader["kolicina"].ToString());
                                g7 = (double.Parse)(reader["vrijednost"].ToString());
                            }

                            //e7 = (reader["Proizvedeno_do"].ToString());
                            ukupk = ukupk + (double.Parse)(reader["kolicina"].ToString());
                        }

                        if (kupac.Contains("SCHWEINFURT"))
                        {
                            if (vo == "TOK")
                            {
                                c8 = (reader["kolicina"].ToString());
                                f8 = (double.Parse)(reader["vrijednost"].ToString());
                            }
                            else
                            {
                                e8 = (reader["kolicina"].ToString());
                                g8 = (double.Parse)(reader["vrijednost"].ToString());
                            }

                            //e7 = (reader["Proizvedeno_do"].ToString());
                            ukupk = ukupk + (double.Parse)(reader["kolicina"].ToString());
                        }

                        if (kupac.Contains("FAG"))
                        {
                            if (vo == "TOK")
                            {
                                c9 = (reader["kolicina"].ToString());
                                f9 = (double.Parse)(reader["vrijednost"].ToString());
                            }
                            else
                            {
                                e9 = (reader["kolicina"].ToString());
                                g9 = (double.Parse)(reader["vrijednost"].ToString());
                            }

                            //e7 = (reader["Proizvedeno_do"].ToString());
                            ukupk = ukupk + (double.Parse)(reader["kolicina"].ToString());
                        }

                        if (kupac.Contains("ROMANIA"))
                        {
                            if (vo == "TOK")
                            {
                                c10 = (reader["kolicina"].ToString());
                                f10 = (double.Parse)(reader["vrijednost"].ToString());
                            }
                            else
                            {
                                e10 = (reader["kolicina"].ToString());
                                g10 = (double.Parse)(reader["vrijednost"].ToString());
                            }

                            //e7 = (reader["Proizvedeno_do"].ToString());
                            ukupk = ukupk + (double.Parse)(reader["kolicina"].ToString());
                        }

                        if (kupac.Contains("Technologies"))  // wupertal
                        {
                            if (vrstap.Contains("Prsten"))  // wupertal
                            {
                                if (vo == "TOK")
                                {
                                    c11 = (reader["kolicina"].ToString());
                                    f11 = (double.Parse)(reader["vrijednost"].ToString());
                                }
                                else
                                {
                                    e11 = (reader["kolicina"].ToString());
                                    g11 = (double.Parse)(reader["vrijednost"].ToString());
                                }

                                //e7 = (reader["Proizvedeno_do"].ToString());
                                ukupk = ukupk + (double.Parse)(reader["kolicina"].ToString());
                            }
                        }

                        if (kupac.Contains("Technologies"))  // wupertal
                        {

                            if (vrstap.Contains("Valjak"))  // wupertal
                            {
                                if (vo == "TOK")
                                {
                                    c12 = (reader["kolicina"].ToString());
                                    f12 = (double.Parse)(reader["vrijednost"].ToString());
                                }
                                else
                                {
                                    e12 = (reader["kolicina"].ToString());
                                    g12 = (double.Parse)(reader["vrijednost"].ToString());
                                }

                                //e7 = (reader["Proizvedeno_do"].ToString());
                                ukupk = ukupk + (double.Parse)(reader["kolicina"].ToString());

                            }
                        }

                        if ((kupac.Contains("KYSUCE"))  && (vrstap.Contains("Ku")))  // 
                            {

                                if (vo == "TOK")
                            {
                                c13 = (reader["kolicina"].ToString());
                                f13 = (double.Parse)(reader["vrijednost"].ToString());
                            }
                            else
                            {
                                e13 = (reader["kolicina"].ToString());
                                g13 = (double.Parse)(reader["vrijednost"].ToString());
                            }

                            //e7 = (reader["Proizvedeno_do"].ToString());
                            ukupk = ukupk + (double.Parse)(reader["kolicina"].ToString());
                        }

                        if ((kupac.Contains("KYSUCE")) && (vrstap.Contains("Prsten")))  // 
                        {
                            if (vo == "TOK")
                            {
                                c17 = (reader["kolicina"].ToString());
                                f17 = (double.Parse)(reader["vrijednost"].ToString());
                            }
                            else
                            {
                                e17 = (reader["kolicina"].ToString());
                                g17 = (double.Parse)(reader["vrijednost"].ToString());
                            }

                            //e7 = (reader["Proizvedeno_do"].ToString());
                            ukupk = ukupk + (double.Parse)(reader["kolicina"].ToString());
                        }


                        if (kupac.Contains("FERRO"))  // wupertal
                        {
                            if (vo == "TOK")
                            {
                                c14 = (reader["kolicina"].ToString());
                                f14 = (double.Parse)(reader["vrijednost"].ToString());
                            }
                            else
                            {
                                e14 = (reader["kolicina"].ToString());
                                g14 = (double.Parse)(reader["vrijednost"].ToString());
                            }

                            //e7 = (reader["Proizvedeno_do"].ToString());
                            ukupk = ukupk + (double.Parse)(reader["kolicina"].ToString());
                        }

                        if (kupac.Contains("SONA"))
                        {
                            if (vo == "TOK")
                            {
                                c15 = (reader["kolicina"].ToString());
                                f15 = (double.Parse)(reader["vrijednost"].ToString());
                            }
                            else
                            {
                                e15 = (reader["kolicina"].ToString());
                                g15 = (double.Parse)(reader["vrijednost"].ToString());
                            }

                            //e7 = (reader["Proizvedeno_do"].ToString());
                            ukupk = ukupk + (double.Parse)(reader["kolicina"].ToString());
                        }

                        if (kupac.Contains("SIGMA"))  // wupertal
                        {
                            if (vo == "TOK")
                            {
                                c16 = (reader["kolicina"].ToString());
                                f16 = (double.Parse)(reader["vrijednost"].ToString());
                            }
                            else
                            {
                                e16 = (reader["kolicina"].ToString());
                                g16 = (double.Parse)(reader["vrijednost"].ToString());
                            }

                            //e7 = (reader["Proizvedeno_do"].ToString());
                            ukupk = ukupk + (double.Parse)(reader["kolicina"].ToString());

                        }

                        if (kupac.Contains("Brasil"))  // Brasil
                        {
                            if (vo == "TOK")
                            {
                                c19 = (reader["kolicina"].ToString());
                                f19 = (double.Parse)(reader["vrijednost"].ToString());
                            }
                            else
                            {
                                e19 = (reader["kolicina"].ToString());
                                g19 = (double.Parse)(reader["vrijednost"].ToString());
                            }

                            //e7 = (reader["Proizvedeno_do"].ToString());
                            ukupk = ukupk + (double.Parse)(reader["kolicina"].ToString());
                        }

                        if (kupac.Contains("TECHNOLO"))  // eltmann
                        {
                            if (vo == "TOK")
                            {
                                c18 = (reader["kolicina"].ToString());
                                f18 = (double.Parse)(reader["vrijednost"].ToString());
                            }
                            else
                            {
                                e18 = (reader["kolicina"].ToString());
                                g18 = (double.Parse)(reader["vrijednost"].ToString());
                            }                            
                            ukupk = ukupk + (double.Parse)(reader["kolicina"].ToString());
                        }


                        if (kupac.Contains("NEUWEG"))  // NSK


                        {
                            if (vo == "TOK")
                            {
                                c21 = (reader["kolicina"].ToString());
                                f21 = (double.Parse)(reader["vrijednost"].ToString());
                            }
                            else
                            {
                                e21 = (reader["kolicina"].ToString());
                                g21 = (double.Parse)(reader["vrijednost"].ToString());
                            }

                            //e7 = (reader["Proizvedeno_do"].ToString());
                            ukupk = ukupk + (double.Parse)(reader["kolicina"].ToString());

                        }
                    }
                }


                string b7 = "", b8 = "", b9 = "", b10 = "", b11 = "", b12 = "", b13 = "", b14 = "", b15 = "", b16 = "", b17 = "", b18 = "", b19 = "", b21 = "";
                string d70 = "", d80 = "", d90 = "", d100 = "", d110 = "", d120 = "", d130 = "", d140 = "", d150 = "", d160 = "", d170 = "", d180 = "", d190 = "", d210 = "";
                string hala = "";
                // kupac, planirano
                using (SqlConnection cn = new SqlConnection(connectionString))
                {
                    cn.Open();
                    //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                    // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                    if (ws <= 2)
                    {
                        sql1 = "rfind.dbo.realizacija '" + dat1 + "','" + dat2 + "',13";
                    }
                    else
                    {
                        sql1 = "rfind.dbo.realizacija '" + dat1 + "','" + dat2 + "',131";
                    }
                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string kupac, vnar = "";

                    while (reader.Read())
                    {

                        if (ws >= 3)
                        {
                            hala = reader["hala"].ToString();
                            if ((ws == 3) && hala.Contains("3")) // pušta samo halu 1
                            {
                                continue;
                            }
                            if ((ws == 4) && hala.Contains("1"))  // pušta samo 3
                            {
                                continue;
                            }
                        }

                        kupac = reader["Kupac"].ToString();
                        vnar = reader["vrstanarudzbe"].ToString();
                        vo = reader["vo"].ToString();
                        if (kupac.Contains("Austria"))
                        {
                            if (vo == "T")
                            {
                                b7 = (reader["norma"].ToString());
                            }
                            else
                            {
                                d70 = (reader["norma"].ToString());
                            }

                        }

                        if (kupac.Contains("SCHWEINFURT"))
                        {

                            if (vo == "T")
                            {
                                b8 = (reader["norma"].ToString());
                            }
                            else
                            {
                                d80 = (reader["norma"].ToString());
                            }

                        }

                        if (kupac.Contains("FAG"))
                        {
                            if (vo == "T")
                            {
                                b9 = (reader["norma"].ToString());
                            }
                            else
                            {
                                d90 = (reader["norma"].ToString());
                            }

                        }

                        if (kupac.Contains("ROMANIA"))
                        {
                            if (vo == "T")
                            {
                                b10 = (reader["norma"].ToString());
                            }
                            else
                            {
                                d100 = (reader["norma"].ToString());
                            }

                        }

                        if (kupac.Contains("Technologies"))  // wupertal
                        {
                            if (vnar.Contains("Prsteni"))  // wupertal
                            {
                                if (vo == "T")
                                {
                                    b11 = (reader["norma"].ToString());
                                }
                                else
                                {
                                    d110 = (reader["norma"].ToString());
                                }

                            }
                        }

                        if (kupac.Contains("Technologies"))  // wupertal
                        {

                            if (vnar.Contains("Valjci"))  // wupertal
                            {
                                if (vo == "T")
                                {
                                    b12 = (reader["norma"].ToString());
                                }
                                else
                                {
                                    d120 = (reader["norma"].ToString());
                                }

                            }
                        }

                        if ((kupac.Contains("KYSUCE")) && (vnar.Contains("Ku")))  // 
                        {
                            if (vo == "T")
                            {
                                b13 = (reader["norma"].ToString());
                            }
                            else
                            {
                                d130 = (reader["norma"].ToString());
                            }

                        }

                        if ((kupac.Contains("KYSUCE")) && (vnar.Contains("Prsten")))  // 
                        {
                            if (vo == "T")
                            {
                                b17 = (reader["norma"].ToString());
                            }
                            else
                            {
                                d170 = (reader["norma"].ToString());
                            }

                        }

                        if (kupac.Contains("FERRO"))  // wupertal
                        {

                            if (vo == "T")
                            {
                                b14 = (reader["norma"].ToString());
                            }
                            else
                            {
                                d140 = (reader["norma"].ToString());
                            }

                        }

                        if (kupac.Contains("SONA"))
                        {
                            if (vo == "T")
                            {
                                b15 = (reader["norma"].ToString());
                            }
                            else
                            {
                                d150 = (reader["norma"].ToString());
                            }

                        }

                        if (kupac.Contains("SIGMA"))  // wupertal
                        {
                            if (vo == "T")
                            {
                                b16 = (reader["norma"].ToString());
                            }
                            else
                            {
                                d160 = (reader["norma"].ToString());
                            }

                        }

                        if (kupac.Contains("Brasil"))  // wupertal
                        {
                            if (vo == "T")
                            {
                                b19 = (reader["norma"].ToString());
                            }
                            else
                            {
                                d190 = (reader["norma"].ToString());
                            }
                        }

                        if (kupac.Contains("TECHNOLOG"))  // ELTMANN
                        {
                            if (vo == "T")
                            {
                                b18 = (reader["norma"].ToString());
                            }
                            else
                            {
                                d180 = (reader["norma"].ToString());
                            }

                        }

                        if (kupac.Contains("NEUWEG"))  // wupertal
                        {
                            if (vo == "T")
                            {
                                b21 = (reader["norma"].ToString());
                            }
                            else
                            {
                                d210 = (reader["norma"].ToString());
                            }

                        }

                    }
                    cn.Close();

                }

                #region             // kupac, broj linija koje ne rade
                int n7 = 0, n8 = 0, n9 = 0, n10 = 0, n11 = 0, n12 = 0, n13 = 0, n14 = 0, n15 = 0, n16 = 0, n17 = 0, n18 = 0, n19 = 0, n20 = 0, n21 = 0, n22=0 ;
                using (SqlConnection cn = new SqlConnection(connectionString))
                {
                    cn.Open();
                    //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                    // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);

                    if (ws<=2)
                        sql1 = "rfind.dbo.ldp_recalc '" + dat1 + "','" + dat2 + "',61";
                    else
                        sql1 = "rfind.dbo.ldp_recalc '" + dat1 + "','" + dat2 + "',610";

                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string kupac, vnar = "";

                    while (reader.Read())
                    {
                        kupac = reader["Kupac"].ToString();
                        vnar = reader["vrstanarudzbe"].ToString();
                        hala = reader["hala"].ToString();
                        if (ws >= 3)
                        {
                            hala = reader["hala"].ToString();
                            if ((ws == 3) && hala.Contains("3")) // pušta samo halu 1
                            {
                                continue;
                            }
                            if ((ws == 4) && hala.Contains("1"))  // pušta samo 3
                            {
                                continue;
                            }
                        }
                        

                        if (kupac.Contains("Austria"))
                        {

                            n7 = (int.Parse)(reader["broj_linija"].ToString());
                        }

                        if (kupac.Contains("SCHWEINFURT"))
                        {

                            n8 = (int.Parse)(reader["broj_linija"].ToString());
                        }

                        if (kupac.Contains("FAG"))
                        {
                            n9 = (int.Parse)(reader["broj_linija"].ToString());
                        }

                        if (kupac.Contains("ROMANIA"))
                        {
                            n10 = (int.Parse)(reader["broj_linija"].ToString());
                        }

                        if (kupac.Contains("Technologies"))  // wupertal
                        {
                            if (vnar.Contains("Prsteni"))  // wupertal
                            {
                                n11 = (int.Parse)(reader["broj_linija"].ToString());  // 
                            }
                        }

                        if (kupac.Contains("Technologies"))  // wupertal
                        {

                            if (vnar.Contains("Valjci"))  // wupertal
                            {
                                n12 = (int.Parse)(reader["broj_linija"].ToString());
                            }
                        }

                        if (kupac.Contains("KYSUCE") && vnar.Contains("Ku"))
                        {
                            n13 = (int.Parse)(reader["broj_linija"].ToString());
                        }

                        if (kupac.Contains("KYSUCE") && vnar.Contains("Prsten"))
                        {
                            n17 = (int.Parse)(reader["broj_linija"].ToString());
                        }

                        if (kupac.Contains("FERRO"))  // wupertal
                        {

                            n14 = (int.Parse)(reader["broj_linija"].ToString());
                        }

                        if (kupac.Contains("SONA"))
                        {
                            n15 = (int.Parse)(reader["broj_linija"].ToString());
                        }

                        if (kupac.Contains("SIGMA"))  // wupertal
                        {
                            n16 = (int.Parse)(reader["broj_linija"].ToString());  //
                        }

                        if (kupac.Contains("TECHNOLOG"))  // eltmann
                        {
                            n18 = (int.Parse)(reader["broj_linija"].ToString());  //
                        }

                        if (kupac.Contains("Brasil"))  // wupertal
                        {
                            n19 = (int.Parse)(reader["broj_linija"].ToString());  //
                        }

                        if (kupac.Contains("NEUWEG"))  // nsk

                        {
                            n21 = (int.Parse)(reader["broj_linija"].ToString());  //
                        }

                        if (kupac=="")
                        {
                      //      n22 = (int.Parse)(reader["broj_linija"].ToString());  //
                        }
                    }
                    cn.Close();
                }

                #endregion  // broj linija koje ne rade

                // kupac, broj linija
                using (SqlConnection cn = new SqlConnection(connectionString))
                {
                    i7 = 0.0; i8 = 0.0; i9 = 0.0; i10 = 0.0; i11 = 0.0; i12 = 0.0; i13 = 0.0; i14 = 0.0; i15 = 0.0; i16 = 0.0; i17=0.0 ; i19 = 0.0;i18 = 0.0; i21 = 0.0;i22 = 0; // 
                    cn.Open();
                    //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                    // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                    if (ws <= 2)
                    {
                        sql1 = "rfind.dbo.realizacija '" + dat1 + "','" + dat2 + "',2";  // za sve hale
                    }
                    else
                    {
                        sql1 = "rfind.dbo.realizacija '" + dat1 + "','" + dat2 + "',20";  // po halama
                    }
                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string kupac, vnar = "";

                    while (reader.Read())
                    {

                        if (ws >= 3)
                        {
                            hala = reader["hala"].ToString();
                            if ((ws == 3) && hala.Contains("3")) // pušta samo halu 1
                            {
                                continue;
                            }
                            if ((ws == 4) && hala.Contains("1"))  // pušta samo 3
                            {
                                continue;
                            }
                        }

                        if (reader["Kupac"] == DBNull.Value)
                        {
                            i22 = i22 + (double.Parse)(reader["broj_linija"].ToString());  //
                            continue;
                        }

                        kupac = reader["Kupac"].ToString();
                        vnar = reader["vrstanarudzbe"].ToString();
                        if (kupac.Contains("Austria"))
                        {
                            i7 = (double.Parse)(reader["broj_linija"].ToString());
                            if (ws==4)
                            {
                                i7 = i7;
                            }

                        }

                        if (kupac.Contains("SCHWEINFURT"))
                        {

                            i8 = (double.Parse)(reader["broj_linija"].ToString());
                        }

                        if (kupac.Contains("FAG"))
                        {
                            i9 = (double.Parse)(reader["broj_linija"].ToString());
                        }

                        if (kupac.Contains("ROMANIA"))
                        {
                            i10 = (double.Parse)(reader["broj_linija"].ToString());
                        }

                        if (kupac.Contains("Technologies"))  // wupertal
                        {
                            if (vnar.Contains("Prsteni"))  // wupertal
                            {
                                i11 = (double.Parse)(reader["broj_linija"].ToString());  // 
                            }
                        }

                        if (kupac.Contains("Technologies"))  // wupertal
                        {

                            if (vnar.Contains("Valjci"))  // wupertal
                            {
                                i12 = (double.Parse)(reader["broj_linija"].ToString());
                            }
                        }

                        if ((kupac.Contains("KYSUCE")) && vnar.Contains("Ku")) // kysuce kućišta
                        {
                            i13 = (double.Parse)(reader["broj_linija"].ToString());
                        }

                        if ((kupac.Contains("KYSUCE")) && vnar.Contains("Prsten")) //
                        {
                            i17 = (double.Parse)(reader["broj_linija"].ToString());
                        }

                        if (kupac.Contains("FERRO"))  // wupertal
                        {
                            i14 = (double.Parse)(reader["broj_linija"].ToString());
                        }

                        if (kupac.Contains("TECHNOLOGIES"))  // ELTMAN 121987
                        {

                            if (vnar.Contains("Valjci"))  // ELTMAN
                            {
                                i18 = (double.Parse)(reader["broj_linija"].ToString());
                            }
                        }



                        if (kupac.Contains("SONA"))
                        {
                            i15 = (double.Parse)(reader["broj_linija"].ToString());
                        }

                        if (kupac.Contains("SIGMA"))  // wupertal
                        {
                            i16 = (double.Parse)(reader["broj_linija"].ToString());  //
                        }

                        if (kupac.Contains("Brasil"))  // wupertal
                        {
                            i19 = (double.Parse)(reader["broj_linija"].ToString());  //
                        }

                        if (kupac.Contains("NEUWEG"))  // nsk
                        {
                            i21 = (double.Parse)(reader["broj_linija"].ToString());  //
                        }
                        

                    }
                    cn.Close();

                }

                
                // kupac, korekcija broja linija
                using (SqlConnection cn = new SqlConnection(connectionString))
                {
                    cn.Open();
                    //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                    // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                    if (ws <= 0)
                    {
                        sql1 = "rfind.dbo.realizacija '" + dat1 + "','" + dat2 + "',22";
                    }
                    else
                    {
                        sql1 = "rfind.dbo.realizacija '" + dat1 + "','" + dat2 + "',220";
                    }
                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string kupac, vnar = "";
                    double ukupno = 0.0;

                    while (reader.Read())
                    {

                        if (ws >= 3)
                        {
                            hala = reader["hala"].ToString();
                            if ((ws == 3) && hala.Contains("3")) // pušta samo halu 1
                            {
                                continue;
                            }
                            if ((ws == 4) && hala.Contains("1"))  // pušta samo 3
                            {
                                continue;
                            }
                        }
                        kupac = reader["Kupac"].ToString();

                        if (kupac.Contains("_Ukupno"))
                        {
                            if (reader["ukupnotr"] != DBNull.Value)
                            {
                                ukupno = (double.Parse)(reader["ukupnotr"].ToString());
                            }
                            else
                            {
                                ukupno = 0.0;
                            }
                        }

                        if (kupac.Contains("Austria"))
                        {

                            i7 = i7 - 1.0 + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                        }

                        if (kupac.Contains("SCHWEINFURT"))
                        {

                            i8 = i8 - 1.0 +  (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                        }

                        if (kupac.Contains("FAG"))
                        {
                            i9 = i9 - 1.0 + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                        }

                        if (kupac.Contains("ROMANIA"))
                        {
                            i10 = i10 - 1.0 +(double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                        }

                        if (kupac.Contains("Technologies"))  // wupertal
                        {

                            i11 = i11 - 1.0 + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                        }
                        
                        if (kupac.Contains("KYSUCE"))  // wupertal
                        {
                            i13 = i13 - 1.0 + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                        }

                        if (kupac.Contains("FERRO"))  // wupertal
                        {

                            i14 = i14 - 1.0 +(double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                        }

                        if (kupac.Contains("SONA"))
                        {
                            i15 = i15 -1.0 + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                        }

                        if (kupac.Contains("SIGMA"))  // wupertal
                        {
                            i16 = i16 - 1 + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                        }

                        if (kupac.Contains("TECHNOLOGIES"))  // eltman
                        {
                            i18 = i18 - 1.0 +(double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                        }

                        if (kupac.Contains("Brasil"))  // wupertal
                        {
                            i19 = i19 - 1.0 + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                        }

                        if (kupac.Contains("NEUWEG"))  // NSK
                        {
                            i21 = i21 - 1.0 +(double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                        }

                    }
                    cn.Close();

                }


                #region                // kupac, lager gotove robe, old

                if (1 == 2)
                {
                    using (SqlConnection cn = new SqlConnection(connectionString))
                    {
                        cn.Open();
                        //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                        // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                        sql1 = "rfind.dbo.realizacija '" + dat1 + "','" + dat2 + "',8";
                        SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                        SqlDataReader reader = sqlCommand.ExecuteReader();
                        string kupac, vnar = "";


                        while (reader.Read())
                        {
                            kupac = reader["nazivpar"].ToString();
                            vnar = reader["vrstanar"].ToString();


                            if (kupac.Contains("Austria"))  // austria
                            {
                                k7 = (double.Parse)(reader["komada"].ToString());
                            }

                            if (kupac.Contains("SCHWEINFURT"))  // 
                            {
                                k8 = (double.Parse)(reader["komada"].ToString());
                            }

                            if (kupac.Contains("FAG"))
                            {
                                k9 = (double.Parse)(reader["komada"].ToString());
                            }

                            if (kupac.Contains("ROMANIA"))
                            {
                                k10 = (double.Parse)(reader["komada"].ToString());
                            }

                            if (kupac.Contains("Technologies"))  // wupertal
                            {
                                if (vnar.Contains("Prsten"))
                                    k11 = (double.Parse)(reader["komada"].ToString());
                                else
                                    k12 = (double.Parse)(reader["komada"].ToString());
                            }

                            if (kupac.Contains("KYSUCE"))  // 
                            {
                                if (vnar.Contains("Ku"))   // kućište
                                    k13 = (double.Parse)(reader["komada"].ToString());
                                else
                                    k17 = (double.Parse)(reader["komada"].ToString());                            
                            }


                            if (kupac.Contains("FERRO"))  // wupertal
                            {
                                k14 = (double.Parse)(reader["komada"].ToString());
                            }

                            if (kupac.Contains("SONA"))
                            {
                                k15 = (double.Parse)(reader["komada"].ToString());
                            }

                            if (kupac.Contains("SIGMA"))  // wupertal
                            {
                                k16 = (double.Parse)(reader["komada"].ToString());
                            }
                            if (kupac.Contains("ELTMAN"))  // wupertal
                            {
                                k18 = (double.Parse)(reader["komada"].ToString());
                            }
                            if (kupac.Contains("Brasil"))  // wupertal
                            {
                                k19 = (double.Parse)(reader["komada"].ToString());
                            }
                            if (kupac.Contains("NEUWEG"))  // wupertal
                            {
                                k21 = (double.Parse)(reader["komada"].ToString());
                            }
                        }
                        cn.Close();
                    }
                }
                #endregion
                // kupac, lager gotove robe, new versandbereite

                using (SqlConnection cn = new SqlConnection(connectionString))
                {
                    cn.Open();
                    //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                    // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                    sql1 = "rfind.dbo.ldp_recalc '0','0',8";
                    sql1 = "rfind.dbo.ldp_recalc '"+dat1+"','"+dat1+"',109";
                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string kupac, vnar = "";

                    while (reader.Read())
                    {
                        kupac = reader["id_par"].ToString();
                        vnar = reader["vrstanar"].ToString();


                        if (kupac == "274")  // austria, Berndorf
                        {
                            k7 = (double.Parse)(reader["komada"].ToString());
                            q7 = (double.Parse)(reader["vrijednost"].ToString());
                        }

                        if (kupac.Contains("121269"))  // SCHWEINFURT
                        {
                            k8 = (double.Parse)(reader["komada"].ToString());
                            q8 = (double.Parse)(reader["vrijednost"].ToString());
                        }

                        if (kupac.Contains("221452"))  // FAG
                        {
                            k9 = (double.Parse)(reader["komada"].ToString());
                            q9 = (double.Parse)(reader["vrijednost"].ToString());
                        }

                        if (kupac.Contains("221453"))  // Romania
                        {
                            k10 = (double.Parse)(reader["komada"].ToString());
                            q10 = (double.Parse)(reader["vrijednost"].ToString());
                        }

                        if (kupac.Contains("121274"))  // wupertal, technologies
                        {
                            if (vnar.Contains("Prsten"))
                            {
                                k11 = (double.Parse)(reader["komada"].ToString());
                                q11 = (double.Parse)(reader["vrijednost"].ToString());
                            }
                            else
                            {
                                k12 = (double.Parse)(reader["komada"].ToString());
                                q12 = (double.Parse)(reader["vrijednost"].ToString());
                            }
                        }

                        if (kupac.Contains("121273"))  // Kysuce
                        {
                            if (vnar.Contains("Ku"))
                            {
                                k13 = (double.Parse)(reader["komada"].ToString());
                                q13 = (double.Parse)(reader["vrijednost"].ToString());
                            }
                            else
                            {
                                k17 = (double.Parse)(reader["komada"].ToString());
                                q17 = (double.Parse)(reader["vrijednost"].ToString());
                            }
                        }

                        if (kupac.Contains("8752"))  // Ferro
                        {
                            k14 = (double.Parse)(reader["komada"].ToString());
                            q14 = (double.Parse)(reader["vrijednost"].ToString());
                        }

                        if (kupac.Contains("121301"))   // Sona
                        {
                            k15 = (double.Parse)(reader["komada"].ToString());
                            q15 = (double.Parse)(reader["vrijednost"].ToString());
                        }

                        if (kupac.Contains("45150"))  // Sigma
                        {
                            k16 = (double.Parse)(reader["komada"].ToString());
                            q16 = (double.Parse)(reader["vrijednost"].ToString());
                        }
                        if (kupac.Contains("121987"))  // Eltman
                        {
                            k18 = (double.Parse)(reader["komada"].ToString());
                            q18 = (double.Parse)(reader["vrijednost"].ToString());
                        }
                        if (kupac.Contains("121278"))  // Brasil
                        {
                            k19 = (double.Parse)(reader["komada"].ToString());
                            q19 = (double.Parse)(reader["vrijednost"].ToString());
                        }
                        if (kupac.Contains("121400"))  // nsk
                        {
                            k21 = (double.Parse)(reader["komada"].ToString());
                            q21 = (double.Parse)(reader["vrijednost"].ToString());
                        }
                    }
                    cn.Close();
                }

                Console.WriteLine("Izračunato stanje Gotove robe trenutno vrijeme " + DateTime.Now);
                // kupac, škart
                double e7tk = 0.00, e7ts = 0.00, post7t = 0.00, e13tk = 0.00, e13ts = 0.00, post13t = 0.00, e8tk = 0.00, e8ts = 0.00, post8t = 0.00, e9tk = 0.00, e9ts = 0.00, post9t = 0.00, e10tk = 0.00, e10ts = 0.00, post10t = 0.00, e21tk = 0.00, e21ts = 0.0, post21t = 0;
                double e11tk = 0.00, e11ts = 0.00, post11t = 0.00, e12tk = 0.00, e12ts = 0.00, post12t = 0.00, e14tk = 0.00, e14ts = 0.00, post14t = 0.00, e15tk = 0.00, e15ts = 0.00, post15t = 0.00, e16tk = 0.00, e16ts = 0.00, post16t = 0.00,e18tk=0.0,e18ts=0.0,e19tk=0.0,post18t=0.0,post18tm=0.0, e19ts = 0.0,  post19t = 0.00;
                double post7tm = 0.00, post8tm = 0.00, post9tm = 0.00, post10tm = 0.00, post11tm = 0.00, post12tm = 0.00, post13tm = 0.00, post14tm = 0.00, post15tm = 0.00, post16tm = 0.00, post19tm = 0.00, post21tm = 0.00;
                double e7to = 0, e8to = 0, e9to = 0, e10to = 0, e11to = 0, e12to = 0, e13to = 0, e14to = 0, e15to = 0, e16to = 0, e17to = 0, e18to = 0, e19to = 0, e21to = 0;
                double e17tk = 0.0, e17ts = 0.0,post17t = 0.0, post17tm = 0.0;post17tm = 0.0;

                // škart

                using (SqlConnection cn = new SqlConnection(connectionString))
                {
                    cn.Open();
                    //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                    // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                    if (ws <= 2)
                    {
                        sql1 = "rfind.dbo.realizacija '" + dat1 + "','" + dat2 + "',3";
                    }
                    else
                    {
                        sql1 = "rfind.dbo.realizacija '" + dat1 + "','" + dat2 + "',31";
                    }
                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string kupac, turm, hala1="";

                    while (reader.Read())
                    {

                        if (ws >= 3)
                        {
                            hala1 = reader["hala"].ToString();
                            if ((ws == 3) && hala1.Contains("3")) // pušta samo halu 1
                            {
                                continue;
                            }
                            if ((ws == 4) && hala1.Contains("1"))  // pušta samo 3
                            {
                                continue;
                            }
                        }

                        kupac = reader["Kupac"].ToString();
                        turm = reader["Turm"].ToString();
                        if (kupac.Contains("Austria"))
                        {
                            if (turm.Contains("Da"))
                            {
                                e7tk = (double.Parse)(reader["kolicinaok"].ToString());
                                e7ts = (double.Parse)(reader["otpadobrada"].ToString());
                                e7to = e7ts;
                                if (e7tk > 0)
                                    post7t = post7t + e7ts / (e7tk + e7ts);

                                e7ts = (double.Parse)(reader["otpadmat"].ToString());

                                if (e7tk > 0)
                                    post7tm = post7tm + e7ts / (e7tk + e7ts);

                            }
                            else
                            {
                                e7tk = (double.Parse)(reader["kolicinaok"].ToString());
                                e7ts = (double.Parse)(reader["otpadobrada"].ToString());
                                e7to = e7ts;
                                if (e7tk > 0)
                                    post7t = post7t + e7ts / (e7tk + e7ts);

                                e7ts = (double.Parse)(reader["otpadmat"].ToString());
                                if (e7tk > 0)
                                    post7tm = post7tm + e7ts / (e7tk + e7ts);

                            }

                        }

                        if (kupac.Contains("SCHWEINFURT"))
                        {

                            if (turm.Contains("Da"))
                            {
                                e8tk = (double.Parse)(reader["kolicinaok"].ToString());
                                e8ts = (double.Parse)(reader["otpadobrada"].ToString());
                                e8to = e8ts;
                                if (e8tk > 0)
                                    post8t = post8t + e8ts / (e8tk + e8ts);

                                e8ts = (double.Parse)(reader["otpadmat"].ToString());
                                if (e8tk > 0)
                                    post8tm = post8tm + e8ts / (e8tk + e8ts);

                            }
                            else
                            {
                                e8tk = (double.Parse)(reader["kolicinaok"].ToString());
                                e8ts = (double.Parse)(reader["otpadobrada"].ToString());
                                e8to = e8ts;
                                if (e8tk > 0)
                                    post8t = post8t + e8ts / (e8tk + e8ts);

                                e8ts = (double.Parse)(reader["otpadmat"].ToString());
                                if (e8tk > 0)
                                    post8tm = post8tm + e8ts / (e8tk + e8ts);

                            }


                        }

                        if (kupac.Contains("FAG"))
                        {
                            if (turm.Contains("Da"))
                            {
                                e9tk = (double.Parse)(reader["kolicinaok"].ToString());
                                e9ts = (double.Parse)(reader["otpadobrada"].ToString());
                                e9to = e9ts;
                                if (e9tk > 0)
                                    post9t = post9t + e9ts / (e9tk + e9ts);

                                e9ts = (double.Parse)(reader["otpadmat"].ToString());
                                if (e9tk > 0)
                                    post9tm = post9tm + e9ts / (e9tk + e9ts);

                            }
                            else
                            {
                                e9tk = (double.Parse)(reader["kolicinaok"].ToString());
                                e9ts = (double.Parse)(reader["otpadobrada"].ToString());
                                e9to = e9ts;
                                if (e9tk > 0)
                                    post9t = post9t + e9ts / (e9tk + e9ts);

                                e9ts = (double.Parse)(reader["otpadmat"].ToString());
                                if (e9tk > 0)
                                    post9tm = post9tm + e9ts / (e9tk + e9ts);

                            }

                        }

                        if (kupac.Contains("ROMANIA"))
                        {
                            if (turm.Contains("Da"))
                            {
                                e10tk = (double.Parse)(reader["kolicinaok"].ToString());
                                e10ts = (double.Parse)(reader["otpadobrada"].ToString());
                                e10to = e10ts;
                                if (e10tk > 0)
                                    post10t = post10t + e10ts / (e10tk + e10ts);

                                e10ts = (double.Parse)(reader["otpadmat"].ToString());
                                if (e10tk > 0)
                                    post10tm = post10tm + e10ts / (e10tk + e10ts);

                            }
                            else
                            {
                                e10tk = (double.Parse)(reader["kolicinaok"].ToString());
                                e10ts = (double.Parse)(reader["otpadobrada"].ToString());
                                e10to = e10ts;
                                if (e10tk > 0)
                                    post10t = post10t + e10ts / (e10tk + e10ts);

                                e10ts = (double.Parse)(reader["otpadmat"].ToString());
                                if (e10tk > 0)
                                    post10tm = post10tm + e10ts / (e10tk + e10ts);

                            }

                        }

                        if (kupac.Contains("Technologies"))  // wupertal
                        {
                            if (kupac.Contains("Prsteni"))  // wupertal
                            {
                                if (turm.Contains("Da"))
                                {
                                    e11tk = (double.Parse)(reader["kolicinaok"].ToString());
                                    e11ts = (double.Parse)(reader["otpadobrada"].ToString());
                                    e11to = e11ts;
                                    if (e11tk > 0)
                                        post11t = post11t + e11ts / (e11tk + e11ts);

                                    e11ts = (double.Parse)(reader["otpadmat"].ToString());
                                    if (e11tk > 0)
                                        post11tm = post11tm + e11ts / (e11tk + e11ts);

                                }
                                else
                                {
                                    e11tk = (double.Parse)(reader["kolicinaok"].ToString());
                                    e11ts = (double.Parse)(reader["otpadobrada"].ToString());
                                    e11to = e11ts;
                                    if (e11tk > 0)
                                        post11t = post11t + e11ts / (e11tk + e11ts);

                                    e11ts = (double.Parse)(reader["otpadmat"].ToString());
                                    if (e11tk > 0)
                                        post11tm = post11tm + e11ts / (e11tk + e11ts);

                                }

                            }
                        }

                        if (kupac.Contains("Technologies"))  // wupertal
                        {

                            if (kupac.Contains("Valjci"))  // wupertal
                            {
                                if (turm.Contains("Da"))
                                {
                                    e12tk = (double.Parse)(reader["kolicinaok"].ToString());
                                    e12ts = (double.Parse)(reader["otpadobrada"].ToString());
                                    e12to = e12ts;
                                    if (e12tk > 0)
                                        post12t = post12t + e12ts / (e12tk + e12ts);

                                    e12ts = (double.Parse)(reader["otpadmat"].ToString());
                                    if (e12tk > 0)
                                        post12tm = post12tm + e12ts / (e12tk + e12ts);

                                }
                                else
                                {
                                    e12tk = (double.Parse)(reader["kolicinaok"].ToString());
                                    e12ts = (double.Parse)(reader["otpadobrada"].ToString());
                                    e12to = e12ts;
                                    if (e12tk > 0)
                                        post12t = post12t + e12ts / (e12tk + e12ts);

                                    e12ts = (double.Parse)(reader["otpadmat"].ToString());
                                    if (e12tk > 0)
                                        post12tm = post12tm + e12ts / (e12tk + e12ts);

                                }

                            }
                        }

                        if ((kupac.Contains("KYSUCE"))  && kupac.Contains("Ku"))
                        {
                            if (turm.Contains("Da"))
                            {
                                e13tk = (double.Parse)(reader["kolicinaok"].ToString());
                                e13ts = (double.Parse)(reader["otpadobrada"].ToString());
                                e13to = e13ts;
                                if (e13tk > 0)
                                    post13t = post13t + e13ts / (e13tk + e13ts);

                                e13ts = (double.Parse)(reader["otpadmat"].ToString());
                                post13tm = post13tm + e13ts / (e13tk + e13ts);

                            }
                            else
                            {
                                e13tk = (double.Parse)(reader["kolicinaok"].ToString());
                                e13ts = (double.Parse)(reader["otpadobrada"].ToString());
                                e13to = e13ts;
                                if (e13tk > 0)
                                    post13t = post13t + e13ts / (e13tk + e13ts);

                                e13ts = (double.Parse)(reader["otpadmat"].ToString());
                                if (e13tk > 0)
                                    post13tm = post13tm + e13ts / (e13tk + e13ts);

                            }

                        }
                        if ((kupac.Contains("KYSUCE")) && kupac.Contains("Prsten"))
                        {
                            if (turm.Contains("Da"))
                            {
                                e17tk = (double.Parse)(reader["kolicinaok"].ToString());
                                e17ts = (double.Parse)(reader["otpadobrada"].ToString());
                                e17to = e17ts;
                                if (e17tk > 0)
                                    post17t = post17t + e17ts / (e17tk + e17ts);

                                e17ts = (double.Parse)(reader["otpadmat"].ToString());
                                post17tm = post17tm + e17ts / (e17tk + e17ts);

                            }
                            else
                            {
                                e17tk = (double.Parse)(reader["kolicinaok"].ToString());
                                e17ts = (double.Parse)(reader["otpadobrada"].ToString());
                                e17to = e17ts;
                                if (e17tk > 0)
                                    post17t = post17t + e17ts / (e17tk + e17ts);

                                e17ts = (double.Parse)(reader["otpadmat"].ToString());
                                if (e17tk > 0)
                                    post17tm = post17tm + e17ts / (e17tk + e17ts);

                            }

                        }
                        
                        if (kupac.Contains("FERRO"))  // wupertal

                        {

                            if (turm.Contains("Da"))
                            {
                                e14tk = (double.Parse)(reader["kolicinaok"].ToString());
                                e14ts = (double.Parse)(reader["otpadobrada"].ToString());
                                e14to = e14ts;
                                if (e14tk > 0)
                                    post14t = post14t + e14ts / (e14tk + e14ts);

                                e14ts = (double.Parse)(reader["otpadmat"].ToString());
                                if (e14tk > 0)
                                    post14tm = post14tm + e14ts / (e14tk + e14ts);

                            }
                            else
                            {
                                e14tk = (double.Parse)(reader["kolicinaok"].ToString());
                                e14ts = (double.Parse)(reader["otpadobrada"].ToString());
                                e14to = e14ts;
                                if (e14tk > 0)
                                    post14t = post14t + e14ts / (e14tk + e14ts);

                                e14ts = (double.Parse)(reader["otpadmat"].ToString());
                                if (e14tk > 0)
                                    post14tm = post14tm + e14ts / (e14tk + e14ts);

                            }

                        }

                        if (kupac.Contains("SONA"))
                        {
                            if (turm.Contains("Da"))
                            {
                                e15tk = (double.Parse)(reader["kolicinaok"].ToString());
                                e15ts = (double.Parse)(reader["otpadobrada"].ToString());
                                e15to = e15ts;
                                if (e15tk > 0)
                                    post15t = post15t + e15ts / (e15tk + e15ts);

                                e15ts = (double.Parse)(reader["otpadmat"].ToString());
                                if (e15tk > 0)
                                    post15tm = post15tm + e15ts / (e15tk + e15ts);

                            }
                            else
                            {
                                e15tk = (double.Parse)(reader["kolicinaok"].ToString());
                                e15ts = (double.Parse)(reader["otpadobrada"].ToString());
                                e15to = e15ts;
                                if (e15tk > 0)
                                    post15t = post15t + e15ts / (e15tk + e15ts);

                                e15ts = (double.Parse)(reader["otpadmat"].ToString());
                                if (e15tk > 0)
                                    post15tm = post15tm + e15ts / (e15tk + e15ts);

                            }

                        }

                        if (kupac.Contains("SIGMA"))  // 
                        {
                            if (turm.Contains("Da"))
                            {
                                e16tk = (double.Parse)(reader["kolicinaok"].ToString());
                                e16ts = (double.Parse)(reader["otpadobrada"].ToString());
                                e16to = e16ts;
                                if (e16tk > 0)
                                    post16t = post16t + e16ts / (e16tk + e16ts);

                                e16ts = (double.Parse)(reader["otpadmat"].ToString());
                                if (e16tk > 0)
                                    post16tm = post16tm + e16ts / (e16tk + e16ts);

                            }
                            else
                            {
                                e16tk = (double.Parse)(reader["kolicinaok"].ToString());
                                e16ts = (double.Parse)(reader["otpadobrada"].ToString());
                                e16to = e16ts;
                                if (e16tk > 0)
                                    post16t = post16t + e16ts / (e16tk + e16ts);

                                e16ts = (double.Parse)(reader["otpadmat"].ToString());
                                if (e16tk > 0)
                                    post16tm = post16tm + e16ts / (e16tk + e16ts);

                            }

                        }


                        if (kupac.Contains("TECHNOLOG"))  // eltmann
                        {
                            if (turm.Contains("Da"))
                            {
                                e18tk = (double.Parse)(reader["kolicinaok"].ToString());
                                e18ts = (double.Parse)(reader["otpadobrada"].ToString());
                                e18to = e18ts;
                                if (e18tk > 0)
                                    post18t = post18t + e18ts / (e18tk + e18ts);

                                e18ts = (double.Parse)(reader["otpadmat"].ToString());
                                if (e18tk > 0)
                                    post18tm = post18tm + e18ts / (e18tk + e18ts);

                            }
                            else
                            {
                                e18tk = (double.Parse)(reader["kolicinaok"].ToString());
                                e18ts = (double.Parse)(reader["otpadobrada"].ToString());
                                e18to = e18ts;
                                if (e18tk > 0)
                                    post18t = post18t + e18ts / (e18tk + e18ts);

                                e18ts = (double.Parse)(reader["otpadmat"].ToString());
                                if (e18tk > 0)
                                    post18tm = post18tm + e18ts / (e18tk + e18ts);

                            }

                        }

                        if (kupac.Contains("Brasil"))  // 
                        {
                            if (turm.Contains("Da"))
                            {
                                e19tk = (double.Parse)(reader["kolicinaok"].ToString());
                                e19ts = (double.Parse)(reader["otpadobrada"].ToString());
                                e19to = e19ts;
                                if (e19tk > 0)
                                    post19t = post19t + e19ts / (e19tk + e19ts);

                                e19ts = (double.Parse)(reader["otpadmat"].ToString());
                                if (e19tk > 0)
                                    post19tm = post19tm + e19ts / (e19tk + e19ts);

                            }
                            else
                            {
                                e19tk = (double.Parse)(reader["kolicinaok"].ToString());
                                e19ts = (double.Parse)(reader["otpadobrada"].ToString());
                                e19to = e19ts;
                                if (e19tk > 0)
                                    post19t = post19t + e19ts / (e19tk + e19ts);

                                e19ts = (double.Parse)(reader["otpadmat"].ToString());
                                if (e19tk > 0)
                                    post19tm = post19tm + e19ts / (e19ts + e19tk);

                            }

                        }

                        if (kupac.Contains("NEUWEG"))  // NSK
                        {
                            if (turm.Contains("Da"))
                            {
                                e21tk = (double.Parse)(reader["kolicinaok"].ToString());
                                e21ts = (double.Parse)(reader["otpadobrada"].ToString());
                                e21to = e21ts;
                                if (e21tk > 0)
                                    post21t = post21t + e21ts / (e21tk + e21ts);

                                e21ts = (double.Parse)(reader["otpadmat"].ToString());
                                if (e21tk > 0)
                                    post21tm = post21tm + e21ts / (e21tk + e21ts);

                            }
                            else
                            {
                                e21tk = (double.Parse)(reader["kolicinaok"].ToString());
                                e21ts = (double.Parse)(reader["otpadobrada"].ToString());
                                e21to = e21ts;
                                if (e21tk > 0)
                                    post21t = post21t + e21ts / (e21tk + e21ts);

                                e21ts = (double.Parse)(reader["otpadmat"].ToString());
                                if (e21tk > 0)
                                    post21tm = post21tm + e21ts / (e21ts + e21tk);

                            }

                        }




                    }
                    cn.Close();

                }

                Console.WriteLine("Izračunat škart   trenutno vrijeme " + DateTime.Now);
                // realizacija ukupno

                //if (1 == 1)
                //{
                double b26 = 0.0, c26 = 0.0, b27 = 0.0, c27 = 0.0;
                using (SqlConnection cn = new SqlConnection(connectionString))
                {
                    cn.Open();
                    //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                    // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                    sql1 = "rfind.dbo.realizacija '" + dat1 + "','" + dat2 + "',5";
                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string pec;

                    while (reader.Read())
                    {
                        //pec = reader["linija"].ToString();

                        //if (pec.Contains("PLANIR"))
                        //{
                        //    b26 = (double.Parse)(reader["K1"].ToString());
                        //    b27 = (double.Parse)(reader["V1"].ToString());

                        //}
                        //else if(pec.Contains("REAL"))
                        //{
                        //    c26 = (double.Parse)(reader["K1"].ToString());
                        //    c27 = (double.Parse)(reader["V1"].ToString());
                        //}
                        if (reader["vrijednost"] != DBNull.Value)
                        {
                            ukupv = (double.Parse)(reader["vrijednost"].ToString());
                            c26 = ukupk;
                            c27 = ukupv;
                        }

                    }
                    cn.Close();
                }
                //}

                Console.WriteLine("Izračunata ukupn. realizacija  trenutno vrijeme " + DateTime.Now);

                // kaliona

                double b31 = 0.0, e31 = 0.0;
                using (SqlConnection cn = new SqlConnection(connectionString))
                {
                    cn.Open();
                    //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                    // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                    sql1 = "rfind.dbo.realizacija '" + dat1 + "','" + dat2 + "',4";
                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string pec;

                    while (reader.Read())
                    {
                        pec = reader["pec"].ToString();
                        if (reader["pec"] != DBNull.Value)
                        {
                            if (pec.Contains("CODERE"))
                            {
                                b31 = (double.Parse)(reader["tezina"].ToString());
                            }
                            else
                            {
                                e31 = (double.Parse)(reader["tezina"].ToString());
                            }
                        }
                    }
                    cn.Close();
                }
                // Kaliona po ldp
                double b312 = 0.0, e312 = 0.0;
                using (SqlConnection cn = new SqlConnection(connectionString))
                {
                    cn.Open();
                    //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                    // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                    sql1 = "rfind.dbo.realizacija '" + dat1 + "','" + dat2 + "',6";
                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string pec;

                    while (reader.Read())
                    {
                        pec = reader["pec"].ToString();
                        if (reader["pec"] != DBNull.Value)
                        {
                            if (pec.Contains("CODERE"))
                            {
                                b312 = (double.Parse)(reader["tezina"].ToString());
                            }
                            else
                            {
                                if (reader["tezina"] != DBNull.Value)
                                {
                                    e312 = (double.Parse)(reader["tezina"].ToString());
                                }
                            }
                        }
                    }
                    cn.Close();
                }
                if (b31 > 0 && e31 < 1)
                {
                    b312 = e312;
                    e312 = 0;
                }


                // broj šarži



                int g31 = 0;
                using (SqlConnection cn = new SqlConnection(connectionString))
                {
                    cn.Open();
                    //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                    // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                    sql1 = "rfind.dbo.realizacija '" + dat1 + "','" + dat2 + "',41";
                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string pec;

                    while (reader.Read())
                    {

                        g31 = g31 + (int.Parse)(reader["brojsarzi"].ToString());

                    }
                    cn.Close();
                }


                double i31 = 0.0;

                using (SqlConnection cn = new SqlConnection(connectionString))
                {
                    cn.Open();
                    //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                    // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                    sql1 = "rfind.dbo.realizacija '" + dat1 + "','" + dat2 + "',6";
                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string pec;

                    while (reader.Read() && (1 == 2))
                    {
                        pec = reader["pec"].ToString();
                        if (reader["pec"] != DBNull.Value)
                        {

                            if (pec.Contains("CODERE"))
                            {
                                b31 = (double.Parse)(reader["tezina"].ToString());
                            }
                            else
                            {
                                e31 = (double.Parse)(reader["tezina"].ToString());
                            }
                        }
                    }
                    cn.Close();
                }

                Console.WriteLine("Izračunate količine kaliona  trenutno vrijeme " + DateTime.Now);

                // kaliona

                double b39 = 0.0, c39 = 0.0, d39 = 0, e39 = 0, f39 = 0, g39 = 0, h39 = 0, i39 = 0, j39 = 0, k39 = 0, l39 = 0, m39 = 0, n39 = 0, o39 = 0, p39 = 0;
                double b42 = 0.0, c42 = 0.0, d42 = 0, e42 = 0, f42 = 0, g42 = 0, h42 = 0, i42 = 0, j42 = 0, k42 = 0, l42 = 0, m42 = 0, n42 = 0, o42 = 0, p42 = 0;
                string idskl = "";

                if (ws < 3)  // samo dnevni i tjedni izvještaj
                {
                    // stanje repromaterijala, količina

                    using (SqlConnection cn = new SqlConnection(connectionString))
                    {
                        cn.Open();
                        //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                        // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                        sql1 = "rfind.dbo.realizacija '" + dat3 + "','" + dat3 + "',7";
                        SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                        SqlDataReader reader = sqlCommand.ExecuteReader();

                        while (reader.Read())
                        {
                            idskl = reader["id_skl"].ToString();

                            switch (idskl)
                            {
                                case "334":
                                    b39 = (double.Parse)(reader["kolicina"].ToString());
                                    break;

                                case "371":
                                    c39 = (double.Parse)(reader["kolicina"].ToString());
                                    break;
                                case "350":
                                    d39 = (double.Parse)(reader["kolicina"].ToString());
                                    break;
                                case "327":
                                    e39 = (double.Parse)(reader["kolicina"].ToString());
                                    break;
                                case "358":
                                    f39 = (double.Parse)(reader["kolicina"].ToString());
                                    break;
                                case "300":
                                    g39 = g39 + (double.Parse)(reader["kolicina"].ToString());
                                    break;
                                case "308":
                                    g39 = g39 + (double.Parse)(reader["kolicina"].ToString());
                                    break;
                                case "390":
                                    h39 = (double.Parse)(reader["kolicina"].ToString());
                                    break;
                                case "361":
                                    i39 = (double.Parse)(reader["kolicina"].ToString());
                                    break;
                                case "394":
                                    j39 = (double.Parse)(reader["kolicina"].ToString());
                                    break;
                                case "392":
                                    k39 = (double.Parse)(reader["kolicina"].ToString());
                                    break;
                                case "252":
                                    l39 = (double.Parse)(reader["kolicina"].ToString());
                                    break;
                                case "153":
                                    m39 = (double.Parse)(reader["kolicina"].ToString());
                                    break;
                                case "155":
                                    n39 = (double.Parse)(reader["kolicina"].ToString());
                                    break;
                                case "118":
                                    o39 = (double.Parse)(reader["tezina"].ToString());
                                    break;
                                case "39":
                                    p39 = (double.Parse)(reader["kolicina"].ToString());
                                    break;
                                default:
                                    break;

                            }


                        }
                        cn.Close();
                    }


                    Console.WriteLine("Izračunate količine materijala  trenutno vrijeme " + DateTime.Now);


                    // stanje repromaterijala, vrijednost

                    using (SqlConnection cn = new SqlConnection(connectionString))
                    {
                        cn.Open();
                        //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                        // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                        sql1 = "rfind.dbo.realizacija '" + dat3 + "','" + dat3 + "',71";
                        SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                        SqlDataReader reader = sqlCommand.ExecuteReader();

                        while (reader.Read())
                        {
                            idskl = reader["id_skl"].ToString();

                            switch (idskl)
                            {
                                case "334":
                                    b42 = (double.Parse)(reader["vrijednost"].ToString());
                                    break;

                                case "371":
                                    c42 = (double.Parse)(reader["vrijednost"].ToString());
                                    break;
                                case "350":
                                    d42 = (double.Parse)(reader["vrijednost"].ToString());
                                    break;
                                case "327":
                                    e42 = (double.Parse)(reader["vrijednost"].ToString());
                                    break;
                                case "358":
                                    f42 = (double.Parse)(reader["vrijednost"].ToString());
                                    break;
                                case "300":
                                    g42 = (double.Parse)(reader["vrijednost"].ToString());
                                    break;
                                case "308":
                                    g42 = (double.Parse)(reader["vrijednost"].ToString());
                                    break;
                                case "390":
                                    h42 = (double.Parse)(reader["vrijednost"].ToString());
                                    break;
                                case "361":
                                    i42 = (double.Parse)(reader["vrijednost"].ToString());
                                    break;
                                case "394":
                                    j42 = (double.Parse)(reader["vrijednost"].ToString());
                                    break;
                                case "392":
                                    k42 = (double.Parse)(reader["vrijednost"].ToString());
                                    break;
                                case "252":
                                    l42 = (double.Parse)(reader["vrijednost"].ToString());
                                    break;
                                case "153":
                                    m42 = (double.Parse)(reader["vrijednost"].ToString());
                                    break;
                                case "155":
                                    n42 = (double.Parse)(reader["vrijednost"].ToString());
                                    break;
                                case "118":
                                    o42 = (double.Parse)(reader["vrijednost"].ToString());
                                    break;
                                case "39":
                                    p42 = (double.Parse)(reader["vrijednost"].ToString());
                                    break;
                                default:
                                    break;

                            }


                        }
                        cn.Close();
                    }


                    Console.WriteLine("Izračunate količine materijala  trenutno vrijeme " + DateTime.Now);
                }

                //Worksheet worksheet = workbook.Worksheets.Item[ws] as Worksheet;
                var worksheet = workbook.Worksheet(ws);
                if (worksheet == null)
                    return;

                int ul24 = 0, ul15 = 0, iz24 = 0, iz15 = 0;
                int naBolovanju = 0, naGO = 0, ostalo = 0;
                if (ws != 2) 
                {

                    // pregled izostanka, zbirno

                    using (SqlConnection cn = new SqlConnection(connectionString))
                    {
                        cn.Open();
                        //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                        // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);

                        sql1 = "rfind.dbo.ldp_recalc '" + dat1 + "','" + dat1 + "',7";
                        SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                        SqlDataReader reader = sqlCommand.ExecuteReader();

                        string vrstao = "";

                        while (reader.Read())
                        {

                            vrstao = reader["vrsta"].ToString();
                            if (vrstao.Contains("Bolovanje"))
                            {
                                naBolovanju = (int.Parse)(reader["broj"].ToString());
                            }

                            if (vrstao.Contains("GO"))
                            {
                                naGO = (int.Parse)(reader["broj"].ToString());
                            }
                            if (vrstao.Contains("Ostalo"))
                            {
                                ostalo = (int.Parse)(reader["broj"].ToString());
                            }
                        }

                        cn.Close();
                    }

                    // koji nisu došli mo mjestu troška
                    using (SqlConnection cn = new SqlConnection(connectionString))
                    {
                        cn.Open();
                        //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                        // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                        if (ws<=2)
                        {
                            sql1 = "rfind.dbo.ldp_recalc '" + datLDP + "','" + datLDP + "',72";
                        }
                        else
                        {
                            sql1 = "rfind.dbo.ldp_recalc '" + datLDP + "','" + datLDP + "',720";
                        }
                        
                        SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                        SqlDataReader reader = sqlCommand.ExecuteReader();
                        //worksizostanci.Cell(4, 3).Value = worksizostanci.Cell(4, 3).Value + " " + datreps;
                        string vrstao = "", mt = "",hala1="";
                        int broj = 0;

                        int i = 0;
                        while (reader.Read())
                        {
                            if (ws >= 3)
                            {
                                hala1 = reader["hala"].ToString();
                                hala1 = reader["hala"].ToString().Trim();
                                if (hala1.Contains("1") || hala1.Contains("3"))
                                {
                                }
                                else
                                {
                                    Console.WriteLine("Hala nije 1,3 ! " + DateTime.Now);
                                    continue;
                                }

                                if ((ws == 3) && hala1.Contains("3")) // pušta samo halu 1
                                {
                                    continue;
                                }
                                if ((ws == 4) && hala1.Contains("1"))  // pušta samo 3
                                {
                                    continue;
                                }
                            }

                            mt = reader["mtroska"].ToString();
                            broj = (int.Parse)(reader["broj"].ToString());
                            vrstao = (reader["vrsta"].ToString());

                            if (mt.Contains("Tokarenje"))
                            {
                                if (vrstao.Contains("Bolovanje"))
                                {
                                    worksheet.Cell(10, 21).Value = broj;
                                }
                                if (vrstao.Contains("Godi"))
                                {
                                    worksheet.Cell(10, 22).Value = broj;
                                }
                                if (vrstao.Contains("Nije"))
                                {
                                    worksheet.Cell(10, 23).Value = broj;
                                }

                            }
                            if (mt.Contains("Kaljenje"))
                            {
                                if (vrstao.Contains("Bolovanje"))
                                {
                                    worksheet.Cell(11, 21).Value = broj;
                                }
                                if (vrstao.Contains("Godi"))
                                {
                                    worksheet.Cell(11, 22).Value = broj;
                                }
                                if (vrstao.Contains("Nije"))
                                {                                    
                                    worksheet.Cell(11, 23).Value = broj;
                                }

                            }
                            if (mt.Contains("Alatnica"))
                            {
                                if (vrstao.Contains("Bolovanje"))
                                {                                    
                                    worksheet.Cell(12, 21).Value = broj;
                                }
                                if (vrstao.Contains("Godi"))
                                {                                    
                                    worksheet.Cell(12, 22).Value = broj;
                                }
                                if (vrstao.Contains("Nije"))
                                {                                    
                                    worksheet.Cell(12, 23).Value = broj;
                                }

                            }
                            if (mt.Contains("Održavanje"))
                            {
                                if (vrstao.Contains("Bolovanje"))
                                {                                    
                                    worksheet.Cell(13, 21).Value = broj;
                                }
                                if (vrstao.Contains("Godi"))
                                {                                    
                                    worksheet.Cell(13, 22).Value = broj;
                                }
                                if (vrstao.Contains("Nije"))
                                {                                    
                                    worksheet.Cell(13, 23).Value = broj;
                                }

                            }
                            if (mt.Contains("Šteleri"))
                            {
                                if (vrstao.Contains("Bolovanje"))
                                {                                    
                                    worksheet.Cell(14, 21).Value = broj;
                                }
                                if (vrstao.Contains("Godi"))
                                {                                    
                                    worksheet.Cell(14, 22).Value = broj;
                                }
                                if (vrstao.Contains("Nije"))
                                {
                                    worksheet.Cell(14, 23).Value = broj;                                    
                                }

                            }
                            if (mt.Contains("Kvaliteta"))
                            {
                                if (vrstao.Contains("Bolovanje"))
                                {                                    
                                    worksheet.Cell(15, 21).Value = broj;
                                }
                                if (vrstao.Contains("Godi"))
                                {
                                    worksheet.Cell(15, 22).Value = broj;
                                }
                                if (vrstao.Contains("Nije"))
                                {
                                    worksheet.Cell(15, 23).Value = broj;
                                }

                            }
                            if (mt.Contains("Zaštita"))
                            {
                                if (vrstao.Contains("Bolovanje"))
                                {
                                    worksheet.Cell(16, 21).Value = broj;
                                }
                                if (vrstao.Contains("Godi"))
                                {
                                    worksheet.Cell(16, 22).Value = broj;
                                }
                                if (vrstao.Contains("Nije"))
                                {
                                    worksheet.Cell(16, 23).Value = broj;
                                }
                            }
                        }
                        i++;

                        cn.Close();
                    }
                    Console.WriteLine("Izostanci po mjestu troška trenutno vrijeme " + DateTime.Now);

                    // broj pristunih po MT

                    using (SqlConnection cn = new SqlConnection(connectionString))
                    {
                        cn.Open();
                        //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                        // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                        if (ws <= 2)
                        {
                            sql1 = "rfind.dbo.ldp_recalc '" + datLDP + "','" + datLDP + "',73";
                        }
                        else
                        {
                            sql1 = "rfind.dbo.ldp_recalc '" + datLDP + "','" + datLDP + "',730";
                        }
                        SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                        SqlDataReader reader = sqlCommand.ExecuteReader();
                        //worksizostanci.Cell(4, 3).Value = worksizostanci.Cell(4, 3).Value + " " + datreps;
                        string vrstao = "", mt = "",hala1="";
                        int broj = 0;

                        int i = 0;
                        while (reader.Read())
                        {
                            if (ws >= 3)
                            {
                                hala1 = reader["hala"].ToString().Trim();
                                if ( hala1.Contains("1") || hala1.Contains("3") )
                                {                                    
                                }
                                else
                                {
                                    Console.WriteLine("Hala nije 1,3 ! " + DateTime.Now);
                                    continue;
                                }

                                if ((ws == 3) && hala1.Contains("3")) // pušta samo halu 1
                                {
                                    continue;
                                }
                                if ((ws == 4) && hala1.Contains("1"))  // pušta samo 3
                                {
                                    continue;
                                }
                            }

                            mt = reader["mtroska"].ToString();
                            broj = (int.Parse)(reader["broj"].ToString());


                            if (mt.Contains("Tokarenje"))
                            {
                                worksheet.Cell(10, 24).Value = broj;

                            }
                            if (mt.Contains("Kaljenje"))
                            {
                                worksheet.Cell(11, 24).Value = broj;
                            }
                            if (mt.Contains("Alatnica"))
                            {
                                worksheet.Cell(12, 24).Value = broj;
                            }
                            if (mt.Contains("Održavanje"))
                            {
                                worksheet.Cell(13, 24).Value = broj;
                            }
                            if (mt.Contains("Šteleri"))
                            {
                                worksheet.Cell(14, 24).Value = broj;
                            }

                            if (mt.Contains("Kvaliteta"))
                            {
                                worksheet.Cell(15, 24).Value = broj;
                            }
                            if (mt.Contains("Zaštita"))
                            {
                                worksheet.Cell(16, 24).Value = broj;
                            }
                            if (mt.Contains("Ukupno"))
                            {
                                worksheet.Cell(7, 19).Value = broj;
                            }

                        }
                        i++;

                        cn.Close();
                    }
                    Console.WriteLine("Došli po mjestu troška trenutno vrijeme " + DateTime.Now);



                } // ws = 1
                  // planiranje share
                //Workbook workbook2 = excel.Workbooks.Open(@"\\192.168.0.3\voditelj pogona\PRIPREMA ZA SASTANAK-PLANIRANJE\Planiranje (SHARE).xlsx", ReadOnly: true, Editable: false);
                // transporti 2017
                //Workbook workbook3 = excel.Workbooks.Open(@"\\192.168.0.3\Referentice$\Transporti 2017.xlsm", ReadOnly: true, Editable: false);
                // Workbook workbook3 = excel.Workbooks.Open(@"c:\brisi\Transporti 2017.xlsx", ReadOnly: true, Editable: false);
                //Workbook workbook2 = excel.Workbooks.Open(@"\\192.168.0.3\voditelj pogona\dokumenti\Planiranje (SHARE)_.xlsx", ReadOnly: true, Editable: false);
                //Worksheet worksheet31 = workbook3.Worksheets.Item[1] as Worksheet;  // 
                // oMissing = null;
                //worksheet31.UsedRange.AutoFilter(1, worksheet31, Microsoft.Office.Interop.Excel.XlAutoFilterOperator.xlAnd, oMissing, false);

                //System.Data.OleDb.OleDbConnection dbConn = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:\\Sample.xlsx;Extended Properties='Excel 12.0 Xml;HDR=YES'");
                //string connectionStringE = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\192.168.0.3\referentice$\Transporti 2017.xlsx;Extended Properties='Excel 12.0 Xml;HDR=YES'";
                string connectionStringE = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=t:\Transporti 2017 .xlsx;Extended Properties='Excel 12.0 Xml;HDR=YES'";
                if (ws < 3 && 1==2)   // samo za dnevni i tjedni izvještaj
                {                  

                    using (System.Data.OleDb.OleDbConnection conne = new System.Data.OleDb.OleDbConnection(connectionStringE))
                    {
                        System.Data.OleDb.OleDbCommand cmd = new System.Data.OleDb.OleDbCommand();
                        cmd.Connection = conne;
                        cmd.CommandText = "SELECT * from [sheet1$] where day(datum)=" + dan1 + " and month(datum)=" + m1 + " and year(datum)=" + g1;

                        conne.Open();
                        System.Data.IDataReader dr = cmd.ExecuteReader();

                        while (dr.Read())
                        {

                            if (dr[12].ToString() == "+") // ulaz
                            {
                                if (dr[10].ToString().Contains("24"))
                                {
                                    ul24 = ul24 + 1;
                                }
                                else
                                {
                                    ul15 = ul15 + 1;

                                }
                            }
                            else
                            {

                                if (dr[10].ToString().Contains("24"))  // izlaz
                                {
                                    iz24 = iz24 + 1;
                                }
                                else
                                {
                                    iz15 = iz15 + 1;

                                }
                            }


                        }
                        conne.Close();
                    }
                }
                //dat10 = "27.7.2017";
                //DateTime vv1;string vv2="",vv3="";

                //Microsoft.Office.Interop.Excel.Range last = worksheet31.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                //int lr = last.Row;
                //Microsoft.Office.Interop.Excel.Range Fruits = excel.get_Range("A1", "A1500");
                //// You should specify all these parameters every time you call this method,
                //// since they can be overridden in the user interface. 
                //string from = "A1";
                //string to = "A1500";

                ////System.Data.OleDb.OleDbConnection dbConn = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:\\Sample.xlsx;Extended Properties='Excel 12.0 Xml;HDR=YES'");
                ////dbConn.Open();
                ////System.Data.OleDb.OleDbDataAdapter db = new System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$]", dbConn);
                ////DataTable dt = new DataTable();
                ////db.Fill(dt);
                ////int iRow = dt.Count;
                ////foreach (DataRow dr in dt.Rows)
                ////    Console.WriteLine(dr[0]);


                ////Range findRng = Fruits.Find("28.7.2017", Missing.Value, XlFindLookIn.xlValues, Missing.Value, Missing.Value, XlSearchDirection.xlNext, false, false, Missing.Value);
                ////vv1 = findRng.Cell(2, 1).Value;
                ////vv2 = findRng.Cell(2, 2).Value;
                ////vv3 = findRng.Cell(2, 13).Value;

                ////findRng = Fruits.FindNext(findRng);
                ////vv1 = findRng.Cell(2, 1).Value;
                ////vv2 = findRng.Cell(2, 2).Value;
                ////vv3 = findRng.Cell(2, 13).Value;

                //Microsoft.Office.Interop.Excel.Range xlFound = Fruits.EntireRow.Find("28.7.2017", Missing.Value, Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues, Microsoft.Office.Interop.Excel.XlLookAt.xlPart, Microsoft.Office.Interop.Excel.XlSearchOrder.xlByColumns, Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext,true, Missing.Value, Missing.Value);
                ////findRng.Select();
                ////int z1 = xlFound.Count;
                ////foreach(Microsoft.Office.Interop.Excel.Range row in findRng.Rows)
                ////{
                ////    vv1 = row.Cell(2, 1).Value;
                ////    vv2 = row.Cell(2, 2).Value;
                ////    vv3 = row.Cell(2, 13).Value;

                ////}

                //Microsoft.Office.Interop.Excel.Range range = worksheet31.UsedRange;

                //DateTime matFilter1 = new DateTime(2017, 7, 27);

                //Microsoft.Office.Interop.Excel.Range jobMatRange = worksheet31.UsedRange; //Define range of Job Material Sheet
                //range.AutoFilter(1, matFilter1, Microsoft.Office.Interop.Excel.XlAutoFilterOperator.xlOr, matFilter1, true); //Apply matFilter1 and matFilter2 to column2 in range

                ////range.AutoFilter(1,dat10 , Microsoft.Office.Interop.Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);


                ////Microsoft.Office.Interop.Excel.Range filteredRange = range.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeVisible);
                //Microsoft.Office.Interop.Excel.Range visibleCells = range.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeVisible, Type.Missing);


                //foreach (Microsoft.Office.Interop.Excel.Range area in visibleCells.Areas)
                //{
                //    vv1 = area.Cell(2, 1).Value;
                //}
                ////Microsoft.Office.Interop.Excel.Range areaRange = filteredRange.Areas.get_Item(1);
                ////DateTime vv1 = areaRange.Cell(2, 1).Value;

                //Worksheet worksheet1 = workbook2.Worksheets.Item[1] as Worksheet;  // BDF
                //Worksheet worksheet2 = workbook2.Worksheets.Item[2] as Worksheet;  // Brasil
                //Worksheet worksheet3 = workbook2.Worksheets.Item[3] as Worksheet;  // Ostalo
                //Worksheet worksheet4 = workbook2.Worksheets.Item[4] as Worksheet;  // Kaliona
                //Worksheet worksheet5 = workbook2.Worksheets.Item[5] as Worksheet;  // pakiranje
                //Worksheet worksheet6 = workbook2.Worksheets.Item[6] as Worksheet;  // FAG
                //Worksheet worksheet7 = workbook2.Worksheets.Item[7] as Worksheet;  // BMW
                //Worksheet worksheet8 = workbook2.Worksheets.Item[8] as Worksheet;  // Kysuce
                //Worksheet worksheet9 = workbook2.Worksheets.Item[9] as Worksheet;  //  Fero preis
                //Worksheet worksheet10 = workbook2.Worksheets.Item[10] as Worksheet;  //  Romania
                //Worksheet worksheet11 = workbook2.Worksheets.Item[11] as Worksheet;  //  eltman
                //Worksheet worksheet12 = workbook2.Worksheets.Item[12] as Worksheet;  //  tecnologies Valjci
                //Worksheet worksheet13 = workbook2.Worksheets.Item[13] as Worksheet;  // tecnologies prsteni
                //Worksheet worksheet14 = workbook2.Worksheets.Item[14] as Worksheet;  //  sona

                if (worksheet == null)
                    return;

                int red1 = 143 + (dana);

                // BDF
                //connectionStringE = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=l:\PRIPREMA ZA SASTANAK-PLANIRANJE\Planiranje (SHARE).xlsx;Extended Properties='Excel 12.0 Xml;HDR=YES'";
                string v1 = "", v2 = "", v3 = "", v4 = "";

                //double v1 = worksheet1.Cell(red1, 8).Value;
                //worksheet.Cell(7, 13).Value = worksheet1.Cell(red1, 8).Value;  //  BDF naš dnevni kapacitet
                //worksheet.Cell(7, 12).Value = worksheet1.Cell(red1, 9).Value;  //  BDF naš dnevni k. trazi
                //worksheet.Cell(7, 11).Value = worksheet1.Cell(red1, 16).Value;  //  BDF  Lager gotove robe  komada

                //double v1 = worksheet1.Cell(red1, 8).Value;

                worksheet.Cell(7, 2).Value = b7;
                worksheet.Cell(7, 4).Value = d70;
                worksheet.Cell(8, 2).Value = b8;
                worksheet.Cell(8, 4).Value = d80;
                worksheet.Cell(9, 2).Value = b9;
                worksheet.Cell(9, 4).Value = d90;
                worksheet.Cell(10, 2).Value = b10;
                worksheet.Cell(10, 4).Value = d100;
                worksheet.Cell(11, 2).Value = b11;
                worksheet.Cell(11, 4).Value = d110;
                worksheet.Cell(12, 2).Value = b12;
                worksheet.Cell(12, 4).Value = d120;
                worksheet.Cell(13, 2).Value = b13;
                worksheet.Cell(13, 4).Value = d130;
                worksheet.Cell(14, 2).Value = b14;
                worksheet.Cell(14, 4).Value = d140;

                worksheet.Cell(15, 2).Value = b15;
                worksheet.Cell(15, 4).Value = d150;

                worksheet.Cell(16, 2).Value = b16;
                worksheet.Cell(16, 4).Value = d160;

                worksheet.Cell(17, 2).Value = b17;
                worksheet.Cell(17, 4).Value = d170;

                worksheet.Cell(18, 2).Value = b18;
                worksheet.Cell(18, 4).Value = d180;

                worksheet.Cell(19, 2).Value = b19;
                worksheet.Cell(19, 4).Value = d190;

                red1 = 141 + dana;

                
                // Pakiranje

                connectionStringE = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=l:\PRIPREMA ZA SASTANAK-PLANIRANJE\Planiranje (SHARE).xlsx;Extended Properties='Excel 12.0 Xml;HDR=YES'";
                //connectionStringE = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=l:\dokumenti\Planiranje (SHARE).xlsx;Extended Properties='Excel 12.0 Xml;HDR=YES'";
                v1 = ""; v2 = ""; v3 = "";
                using (System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection(connectionStringE))
                {
                    System.Data.OleDb.OleDbCommand cmd = new System.Data.OleDb.OleDbCommand();
                    cmd.Connection = conn;
                    cmd.CommandText = "SELECT * from [pakiranje$] where day(datum)=" + dan1 + " and month(datum)=" + m1 + " and year(datum)=" + g1;

                      conn.Open();
                    System.Data.IDataReader dr = cmd.ExecuteReader();

                    while (dr.Read())
                    {

                        v1 = dr[4].ToString();
                    }
                }

                worksheet.Cell(23, 2).Value = v1;  // FAG
                v1 = ""; v2 = ""; v3 = b9; v4 = "";

                // brasil
                //red1 = 143 + dana;
                //v1 = ""; v2 = ""; v3 = ""; v4 = "";
                //using (System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection(connectionStringE))
                //{
                //    System.Data.OleDb.OleDbCommand cmd = new System.Data.OleDb.OleDbCommand();
                //    cmd.Connection = conn;
                //    cmd.CommandText = "SELECT * from [Schaeffler Brasil Ltda.$] where day(datum)=" + dan1 + " and month(datum)=" + m1 + " and year(datum)=" + g1;

                //    conn.Open();
                //    System.Data.IDataReader dr = cmd.ExecuteReader();

                //    while (dr.Read())
                //    {

                //        v1 = dr[7].ToString(); // kupac trazi
                //        v2 = dr[6].ToString(); // naš kapacitet
                //        v3 = dr[1].ToString();
                //        v4 = dr[12].ToString();
                //    }
                //}

                //v4 = k15.ToString();
                //worksheet.Cell(15, 12).Value = v1;  //  sona kupac trazi
                //worksheet.Cell(15, 13).Value = v2;  //  sona  naš kapacitet
                //worksheet.Cell(15, 2).Value = v3;  //  sona planirano
                //worksheet.Cell(15, 11).Value = v4;  //  sona LGR kom.


                red1 = 141 + dana;
                v1 = ""; v2 = ""; v3 = ""; v4 = "";
                if (ws < 3)  // samo dnevni i tjedni izvještaj
                { 
                    using (System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection(connectionStringE))
                    {
                        System.Data.OleDb.OleDbCommand cmd = new System.Data.OleDb.OleDbCommand();
                        cmd.Connection = conn;
                        cmd.CommandText = "SELECT * from [KALIONA$] where day(datum)=" + dan1 + " and month(datum)=" + m1 + " and year(datum)=" + g1;

                        conn.Open();
                        System.Data.IDataReader dr = cmd.ExecuteReader();

                        while (dr.Read())
                        {

                            v1 = dr[17].ToString(); // kupac trazi

                        }
                        conn.Close();
                    }
                }
                if (ws == 1)
                    worksheet.Cell(33, 9).Value = v1;  //  Kaliona   cementacija kg


                //workbook2.Close(false);
                //workbook2.Close(false, Type.Missing, Type.Missing);
                //workbook2.Close(false, Type.Missing, Type.Missing);

                GC.Collect();


                //objExcel.Application.Quit();
                //objExcel.Quit();
                Console.WriteLine("Zatvoreno Planiranje share   trenutno vrijeme " + DateTime.Now);

                // }  // if ws==1
                //Range row1 = worksheet.Cell(7, 3];
                //Range row2 = worksheet.Cell(13, 3];

                //row1.Value = c7;
                //row2.Value = c13;

                
                worksheet.Cell(1, 1).Value = worksheet.Cell(1, 1).Value + " " + dat10;

                worksheet.Cell(16, 2).Value = b16;  // planiranje  Sigma
                //worksheet.Cell(16, 11).Value = k16;  // LGR  Sigma

                worksheet.Cell(19, 2).Value = b19;  // planiranje Brasil
                worksheet.Cell(21, 2).Value = b21;  // planiranje nsk

                worksheet.Cell(7, 3).Value = c7;  // količina tok
                worksheet.Cell(8, 3).Value = c8;
                worksheet.Cell(9, 3).Value = c9;
                worksheet.Cell(10, 3).Value = c10;
                worksheet.Cell(11, 3).Value = c11;
                worksheet.Cell(12, 3).Value = c12;
                worksheet.Cell(13, 3).Value = c13;
                worksheet.Cell(14, 3).Value = c14;
                worksheet.Cell(15, 3).Value = c15;
                worksheet.Cell(16, 3).Value = c16;
                worksheet.Cell(17, 3).Value = c17;
                worksheet.Cell(18, 3).Value = c18;
                worksheet.Cell(19, 3).Value = c19;
                worksheet.Cell(21, 3).Value = c21;

                worksheet.Cell(7, 5).Value = e7;  // količina, dodatne operacije
                worksheet.Cell(8, 5).Value = e8;
                worksheet.Cell(9, 5).Value = e9;
                worksheet.Cell(10, 5).Value = e10;
                worksheet.Cell(11, 5).Value = e11;
                worksheet.Cell(12, 5).Value = e12;
                worksheet.Cell(13, 5).Value = e13;
                worksheet.Cell(14, 5).Value = e14;
                worksheet.Cell(15, 5).Value = e15;
                worksheet.Cell(16, 5).Value = e16;
                worksheet.Cell(17, 5).Value = e17;
                worksheet.Cell(18, 5).Value = e18;
                worksheet.Cell(19, 5).Value = e19;
                worksheet.Cell(21, 5).Value = e21;

                worksheet.Cell(7, 6).Value = f7;  // vrijednost tokarenja
                worksheet.Cell(8, 6).Value = f8;
                worksheet.Cell(9, 6).Value = f9;
                worksheet.Cell(10, 6).Value = f10;
                worksheet.Cell(11, 6).Value = f11;
                worksheet.Cell(12, 6).Value = f12;
                worksheet.Cell(13, 6).Value = f13;
                worksheet.Cell(14, 6).Value = f14;
                worksheet.Cell(15, 6).Value = f15;
                worksheet.Cell(16, 6).Value = f16;
                worksheet.Cell(17, 6).Value = f17;
                worksheet.Cell(18, 6).Value = f18;
                worksheet.Cell(19, 6).Value = f19;
                worksheet.Cell(21, 6).Value = f21;

                
                worksheet.Cell(7, 7).Value = g7;  // vrijednost dodatne operacije
                worksheet.Cell(8, 7).Value = g8;
                worksheet.Cell(9, 7).Value = g9;
                worksheet.Cell(10, 7).Value = g10;
                worksheet.Cell(11, 7).Value = g11;
                worksheet.Cell(12, 7).Value = g12;
                worksheet.Cell(13, 7).Value = g13;
                worksheet.Cell(14, 7).Value = g14;
                worksheet.Cell(15, 7).Value = g15;
                worksheet.Cell(16, 7).Value = g16;
                worksheet.Cell(17, 7).Value = g17;
                worksheet.Cell(18, 7).Value = g18;
                worksheet.Cell(19, 7).Value = g19;
                worksheet.Cell(21, 7).Value = g21;

                worksheet.Cell(7, 13).Value = i7;   // broj linija
                worksheet.Cell(8, 13).Value = i8;
                worksheet.Cell(9, 13).Value = i9;
                worksheet.Cell(10, 13).Value = i10;
                worksheet.Cell(11, 13).Value = i11;
                worksheet.Cell(12, 13).Value = i12;
                worksheet.Cell(13, 13).Value = i13;
                worksheet.Cell(14, 13).Value = i14;
                worksheet.Cell(15, 13).Value = i15;
                worksheet.Cell(16, 13).Value = i16;
                worksheet.Cell(17, 13).Value = i17;
                worksheet.Cell(18, 13).Value = i18;
                worksheet.Cell(19, 13).Value = i19;
                worksheet.Cell(21, 13).Value = i21;
                worksheet.Cell(22, 13).Value = i22;

                worksheet.Cell(7, 15).Value = j7;   // broj štelanja
                worksheet.Cell(8, 15).Value = j8;
                worksheet.Cell(9, 15).Value = j9;
                worksheet.Cell(10, 15).Value = j10;
                worksheet.Cell(11, 15).Value = j11;
                worksheet.Cell(12, 15).Value = j12;
                worksheet.Cell(13, 15).Value = j13;
                worksheet.Cell(14, 15).Value = j14;
                worksheet.Cell(15, 15).Value = j15;
                worksheet.Cell(16, 15).Value = j16;
                worksheet.Cell(17, 15).Value = j17;
                worksheet.Cell(18, 15).Value = j18;
                worksheet.Cell(19, 15).Value = j19;
                worksheet.Cell(21, 15).Value = j21;

                if (ws <= 2)
                {

                    if (ws == 1)
                    {
                        worksheet.Cell(7, 16).Value = k7;   // LGR, komada
                        worksheet.Cell(8, 16).Value = k8;
                        worksheet.Cell(9, 16).Value = k9;
                        worksheet.Cell(10, 16).Value = k10;
                        worksheet.Cell(11, 16).Value = k11;
                        worksheet.Cell(12, 16).Value = k12;
                        worksheet.Cell(13, 16).Value = k13;
                        worksheet.Cell(14, 16).Value = k14;
                        worksheet.Cell(15, 16).Value = k15;
                        worksheet.Cell(16, 16).Value = k16;
                        worksheet.Cell(17, 16).Value = k17;
                        worksheet.Cell(18, 16).Value = k18;
                        worksheet.Cell(19, 16).Value = k19;
                        worksheet.Cell(21, 16).Value = k21;

                        worksheet.Cell(7, 17).Value = q7;   // LGR, vrijednost
                        worksheet.Cell(8, 17).Value = q8;
                        worksheet.Cell(9, 17).Value = q9;
                        worksheet.Cell(10, 17).Value = q10;
                        worksheet.Cell(11, 17).Value = q11;
                        worksheet.Cell(12, 17).Value = q12;
                        worksheet.Cell(13, 17).Value = q13;
                        worksheet.Cell(14, 17).Value = q14;
                        worksheet.Cell(15, 17).Value = q15;
                        worksheet.Cell(16, 17).Value = q16;
                        worksheet.Cell(17, 17).Value = q17;
                        worksheet.Cell(18, 17).Value = q18;
                        worksheet.Cell(19, 17).Value = q19;
                        worksheet.Cell(21, 17).Value = q21;



                    }
                    


                }

                worksheet.Cell(7, 14).Value = n7;   // broj linija koje ne rade
                worksheet.Cell(8, 14).Value = n8;
                worksheet.Cell(9, 14).Value = n9;
                worksheet.Cell(10, 14).Value = n10;
                worksheet.Cell(11, 14).Value = n11;
                worksheet.Cell(12, 14).Value = n12;
                worksheet.Cell(13, 14).Value = n13;
                worksheet.Cell(14, 14).Value = n14;
                worksheet.Cell(15, 14).Value = n15;
                worksheet.Cell(16, 14).Value = n16;
                worksheet.Cell(17, 14).Value = n17;
                worksheet.Cell(18, 14).Value = n18;
                worksheet.Cell(19, 14).Value = n19;
                worksheet.Cell(21, 14).Value = n21;
                worksheet.Cell(22, 14).Value = n22;    // linija nije  zaplanirana


                // škart postoci
                worksheet.Cell(7, 9).Value = Math.Round(post7t, 4);   // škart obrade
                worksheet.Cell(8, 9).Value = Math.Round(post8t, 4);
                worksheet.Cell(9, 9).Value = Math.Round(post9t, 4);
                worksheet.Cell(10, 9).Value = Math.Round(post10t, 4);
                worksheet.Cell(11, 9).Value = Math.Round(post11t, 4);
                worksheet.Cell(12, 9).Value = Math.Round(post12t, 4);
                double posto13 = Math.Round(post13t, 4);
                worksheet.Cell(13, 9).Value = posto13;
                worksheet.Cell(14, 9).Value = Math.Round(post14t, 4);
                worksheet.Cell(15, 9).Value = Math.Round(post15t, 4);
                worksheet.Cell(16, 9).Value = Math.Round(post16t, 4);
                worksheet.Cell(17, 9).Value = Math.Round(post17t, 4);
                worksheet.Cell(18, 9).Value = Math.Round(post18t, 4);
                worksheet.Cell(19, 9).Value = Math.Round(post19t, 4);
                worksheet.Cell(21, 9).Value = Math.Round(post21t, 4);
                double pos131 = Math.Round(post13tm, 4);


                worksheet.Cell(7, 11).Value = Math.Round(post7tm, 4);   // škart materijala
                worksheet.Cell(8, 11).Value = Math.Round(post8tm, 4);
                worksheet.Cell(9, 11).Value = Math.Round(post9tm, 4);
                worksheet.Cell(9, 11).Value = Math.Round(post10tm, 4);
                worksheet.Cell(11, 11).Value = Math.Round(post11tm, 4);
                worksheet.Cell(12, 11).Value = Math.Round(post12tm, 4);
                worksheet.Cell(13, 11).Value = Math.Round(post13tm, 4);
                worksheet.Cell(14, 11).Value = Math.Round(post14tm, 4);
                worksheet.Cell(15, 11).Value = Math.Round(post15tm, 4);
                worksheet.Cell(16, 11).Value = Math.Round(post16tm, 4);
                worksheet.Cell(17, 11).Value = Math.Round(post17tm, 4);
                worksheet.Cell(18, 11).Value = Math.Round(post18tm, 4);
                worksheet.Cell(19, 11).Value = Math.Round(post19tm, 4);
                worksheet.Cell(21, 11).Value = Math.Round(post21tm, 4);

                // škart komadi
                worksheet.Cell(7, 10).Value = e7to;   // škart obrade
                worksheet.Cell(8, 10).Value = e8to;
                worksheet.Cell(9, 10).Value = e9to;
                worksheet.Cell(10, 10).Value = e10to;
                worksheet.Cell(11, 10).Value = e11to;
                worksheet.Cell(12, 10).Value = e12to;

                worksheet.Cell(13, 10).Value = e13to;
                worksheet.Cell(14, 10).Value = e14to;
                worksheet.Cell(15, 10).Value = e15to;
                worksheet.Cell(16, 10).Value = e16to;
                worksheet.Cell(17, 10).Value = e16to;
                worksheet.Cell(18, 10).Value = e18to;
                worksheet.Cell(19, 10).Value = e19to;
                worksheet.Cell(21, 10).Value = e21to;

                worksheet.Cell(7, 12).Value = e7ts;   // škart materijala
                worksheet.Cell(8, 12).Value = e8ts;
                worksheet.Cell(9, 12).Value = e9ts;
                worksheet.Cell(10, 12).Value = e10ts;
                worksheet.Cell(11, 12).Value = e11ts;
                worksheet.Cell(12, 12).Value = e12ts;
                worksheet.Cell(13, 12).Value = e13ts;
                worksheet.Cell(14, 12).Value = e14ts;
                worksheet.Cell(15, 12).Value = e15ts;
                worksheet.Cell(16, 12).Value = e16ts;
                worksheet.Cell(17, 12).Value = e17ts;
                worksheet.Cell(17, 10).Value = e18ts;
                worksheet.Cell(19, 12).Value = e19ts;
                worksheet.Cell(21, 12).Value = e21ts;

                int kol1 = (int.Parse)(worksheet.Cell(24, 2).Value.ToString());
                int kol2 = (int.Parse)(worksheet.Cell(24, 4).Value.ToString());

                //            worksheet.Cell(26, 2).Value = b26;    // Realizacija ukupno: planirano
                worksheet.Cell(28, 2).Value = kol1+kol2;

                kol1 = (int.Parse)(worksheet.Cell(24, 3).Value.ToString());
                kol2 = (int.Parse)(worksheet.Cell(24, 5).Value.ToString());

                worksheet.Cell(28, 3).Value = kol1 + kol2;     // ukupna količina
                string s1 = worksheet.Cell(24, 6).Value.ToString(); //.Replace(',', '.');
                decimal kol21 = Convert.ToDecimal((worksheet.Cell(24, 6).Value.ToString()));
                decimal kol22 =  Convert.ToDecimal((worksheet.Cell(24, 7).Value.ToString()));

                //           worksheet.Cell(27, 2).Value = b27;    // Realizacija ukupno: realizirano
                worksheet.Cell(29, 3).Value = kol21 + kol22;     // ukupna količinac27;    // ukupna vrijednost
                

                if (ws == 2)
                {
                    worksheet.Cell(33, 2).Value = b312;    // kaliona
                    worksheet.Cell(33, 5).Value = e312;    // kaliona
                    worksheet.Cell(33, 9).Value = (e312 + b312) * 0.33;  // in
                    worksheet.Cell(33, 7).Value = g31;   // broj šarži
                    worksheet.Cell(33, 10).Value = (e312 + b312);   // in   // broj šarži
                }


                if (ws == 1)
                {
                    worksheet.Cell(33, 3).Value = b31;    // kaliona
                    worksheet.Cell(33, 6).Value = e31;
                    worksheet.Cell(33, 8).Value = g31;   // broj šarži
                    worksheet.Cell(33, 10).Value = (e31 + b31);   // in
                    worksheet.Cell(33, 11).Value = (e31 + b31) * 0.33;  // in
                    worksheet.Cell(34, 3).Value = b312;    // kaliona
                    worksheet.Cell(34, 6).Value = e312;
                    //                worksheet.Cell(34, 7).Value = g31;   // broj šarži
                    worksheet.Cell(34, 10).Value = (e312 + b312);
                    worksheet.Cell(34, 11).Value = (e312 + b312) * 0.33;  // out
                    worksheet.Cell(2, 15).Value = worksheet.Cell(2, 15).Value + " " + dat10;
                    //  worksheet.Cell(2, 20).Value = worksheet.Cell(2, 20).Value + " " + dat10;
                    //  worksheet.Cell(9, 20).Value = dat10;

                    worksheet.Cell(41, 1).Value = dat30;

                    //worksheet.Cell(36, 1).Value = worksheet.Cell(34, 1).Value + " " + dat10;
                    worksheet.Cell(28, 6).Value = ul24;    // Transporti 24 t
                    worksheet.Cell(28, 7).Value = iz24;
                    worksheet.Cell(29, 6).Value = ul15;    // Transporti 1.5T
                    worksheet.Cell(29, 7).Value = iz15;

                    worksheet.Cell(42, 2).Value = b39;   // količina materijala, HAY
                    worksheet.Cell(42, 3).Value = c39;   // količina materijala
                    worksheet.Cell(42, 4).Value = d39;   // količina materijala
                    worksheet.Cell(42, 5).Value = e39;   // količina materijala
                    worksheet.Cell(42, 6).Value = f39;   // količina materijala
                    worksheet.Cell(42, 7).Value = g39;   // količina materijala
                    worksheet.Cell(42, 8).Value = h39;   // količina materijala
                    worksheet.Cell(42, 9).Value = i39;   // količina materijala
                    worksheet.Cell(42, 10).Value = j39;   // količina materijala
                    worksheet.Cell(42, 11).Value = k39;   // količina materijala
                    worksheet.Cell(42, 12).Value = l39;   // količina materijala
                    worksheet.Cell(42, 13).Value = m39;   // količina materijala
                    worksheet.Cell(42, 14).Value = n39;   // količina materijala
                    worksheet.Cell(42, 15).Value = o39;   // količina materijala
                    worksheet.Cell(42, 16).Value = p39;   // količina materijala

                    worksheet.Cell(43, 2).Value = b42;   // vrijednost materijala, HAY
                    worksheet.Cell(43, 3).Value = c42;   // vrijednost materijala
                    worksheet.Cell(43, 4).Value = d42;   // vrijednost materijala
                    worksheet.Cell(43, 5).Value = e42;   // vrijednost materijala
                    worksheet.Cell(43, 6).Value = f42;   // vrijednost materijala
                    worksheet.Cell(43, 7).Value = g42;   // vrijednost materijala
                    worksheet.Cell(43, 8).Value = h42;   // vrijednost materijala
                    worksheet.Cell(43, 9).Value = i42;   // vrijednost materijala
                    worksheet.Cell(43, 10).Value = j42;   // vrijednost materijala
                    worksheet.Cell(43, 11).Value = k42;   // vrijednost materijala
                    worksheet.Cell(43, 12).Value = l42;   // vrijednost materijala
                    worksheet.Cell(43, 13).Value = m42;   // vrijednost materijala
                    worksheet.Cell(43, 14).Value = n42;   // vrijednost materijala
                    worksheet.Cell(43, 15).Value = o42;   // vrijednost materijala
                    worksheet.Cell(43, 16).Value = p42;   // vrijednost materijala

                    //worksheet.Cell(7, 20).Value = naBolovanju;   // Izostanci, na bolovanju
                    //worksheet.Cell(7, 21).Value = naGO;   // na godišnjem
                    //worksheet.Cell(7, 22).Value = ostalo;   // ostalo
                    //worksheet.Cell(8, 22).Value = ostalo+naGO+naBolovanju ;   // ukupno
                }
                Console.WriteLine("Napunjen dsrddmmyyyy   trenutno vrijeme " + DateTime.Now);
                if (worksheet == null)
                    return;


            }  // for 


            if (1 == 2)
            {
                //var fi = new FileInfo(fileName);
                //if (fi.Exists) File.Delete(fileName);

                //excel.Application.ActiveWorkbook.SaveAs(fileName);
                //Console.WriteLine("Snimljen file ddmmyyyy   trenutno vrijeme " + DateTime.Now);
                ////Console.ReadKey();
                ////workbook.Close(false);
                ////excel.Application.Quit();
                ////excel.Quit()
                ////workbook.Close(true, Type.Missing, Type.Missing);
                //workbook.Close(false, Type.Missing, Type.Missing);
                //excel.Application.Quit();
                //excel.Quit();

                //excel = new Application();
                //app = new Microsoft.Office.Interop.Excel.Application();
                //workbook = excel.Workbooks.Open(fileName, ReadOnly: false, Editable: true);
            }

            //Worksheet worksheet3r = workbook.Worksheets.Item[2] as Worksheet;
            // 21.03
            //Worksheet worksheet3r2 = workbook.Worksheets.Item[5] as Worksheet;   // ldp po radnicima
            //Worksheet worksheet4l = workbook.Worksheets.Item[6] as Worksheet;    // ldp po linijama

            //Worksheet workshstelanje = workbook.Worksheets.Item[7] as Worksheet;    // Pregled štelanja,aktivnosti
            //Worksheet worksizostanci = workbook.Worksheets.Item[8] as Worksheet;    // Pregled izostanaka
           
            var worksheet3r2 = workbook.Worksheet(5);   // ldp po radnicima
            var worksheet4l  = workbook.Worksheet(6);    // ldp po linijama

            var workshstelanje = workbook.Worksheet(7);    // Pregled štelanja,aktivnosti
            var worksizostanci = workbook.Worksheet(8);    // Pregled izostanaka


            // LDP
            // radnici

            worksheet3r2.Cell(1, 2).Value = worksheet3r2.Cell(1, 2).Value + " " + datLDPTitle;
            //            worksheet3r.Cell(2, 6).Value = smjenanaziv;
            // linije
            worksheet4l.Cell(1, 2).Value = worksheet4l.Cell(1, 2).Value + " " + datLDPTitle;
            //          worksheet4l.Cell(2, 6).Value = smjenanaziv;


            // pregled izostanka, popis imena,mt

            using (SqlConnection cn = new SqlConnection(connectionString))
            {
                cn.Open();
                //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                sql1 = "rfind.dbo.ldp_recalc '" + datLDP + "','" + datLDP + "',71";
                SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                SqlDataReader reader = sqlCommand.ExecuteReader();
                worksizostanci.Cell(4, 3).Value = worksizostanci.Cell(4, 3).Value + " " + datreps;
                string vrstao = "";
                int i = 0;
                while (reader.Read())
                {
                    if (1 == 1)
                    {

                        worksizostanci.Cell(8 + i, 1).Value = reader["mtroska"];
                        worksizostanci.Cell(8 + i, 2).Value = reader["hala"];
                        worksizostanci.Cell(8 + i, 3).Value = reader["smjena"];
                        worksizostanci.Cell(8 + i, 4).Value = reader["linija"].ToString();
                        worksizostanci.Cell(8 + i, 5).Value = reader["prezime"].ToString();
                        worksizostanci.Cell(8 + i, 6).Value = reader["ime"].ToString();
                        worksizostanci.Cell(8 + i, 7).Value = reader["vrsta"].ToString();
                        if (i % 2==0)
                            worksizostanci.Range(8 + i,1,8+i,7).Style.Fill.BackgroundColor = XLColor.WhiteSmoke;
                        else
                            worksizostanci.Range(8 + i, 1, 8 + i, 7).Style.Fill.BackgroundColor = XLColor.White;

                    }
                    i++;

                }

                cn.Close();
            }
            Console.WriteLine("Izostanci  trenutno vrijeme " + DateTime.Now);

            string smj = "4";
            string dat13 = dat1 + " 06:00:00";
            string dat23 = dat3 + " 06:00:00";
            dat13 = datLDP;
            dat23 = dat3;


            using (SqlConnection cn = new SqlConnection(connectionString))
            {
                cn.Open();
                //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                sql1 = "rfind.dbo.ldp_recalc '" + dat13 + "','" + dat23 + "'," + smj; ;
                int k1 = 0;
                int n1 = 0;

                if (1 == 2)
                {// prvi ldp po radncima
                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string pec;
                    int i = 0;
                    int ukupn = 0, ukupp = 0, ukupkol = 0;

                    while (reader.Read())
                    {

                        //    worksheet3r.Cell(5 + i, 1).Value = reader["datum"];
                        //    worksheet3r.Cell(5 + i, 2).Value = (reader["hala"].ToString());
                        //    worksheet3r.Cell(5 + i, 3).Value = (reader["linija"].ToString());
                        //    worksheet3r.Cell(5 + i, 4).Value = (reader["nazivpro"].ToString());
                        //    worksheet3r.Cell(5 + i, 5).Value = (reader["radnik"].ToString());
                        //    worksheet3r.Cell(5 + i, 6).Value = reader["norma"];
                        //    worksheet3r.Cell(5 + i, 7).Value = reader["planirano"];
                        //    worksheet3r.Cell(5 + i, 8).Value = (reader["kolicina"].ToString());
                        //    worksheet3r.Cell(5 + i, 9).Value = (reader["kupac"].ToString());
                        //    worksheet3r.Cell(5 + i, 10).Value = (reader["napomena"].ToString());
                        //    worksheet3r.Cell(5 + i, 11).Value = (reader["aktivnost"].ToString());

                        //    if ((int.Parse)(reader["kolicina"].ToString()) > (int.Parse)(reader["planirano"].ToString()))
                        //        worksheet3r.Cell(5 + i, 8].Interior.Color = XlRgbColor.rgbGreenYellow;

                        //    ukupn = ukupn + (int.Parse)(reader["norma"].ToString());

                        //    if (reader["planirano"] != DBNull.Value)
                        //    {
                        //        ukupp = ukupp + (int.Parse)(reader["planirano"].ToString());
                        //        ukupkol = ukupkol + (int.Parse)(reader["kolicina"].ToString());
                        //    }

                        //    i++;

                    }
                    //worksheet3r.Cell(5 + i, 5).Value = "Ukupno :";
                    //worksheet3r.Cell(5 + i, 6).Value = ukupn;
                    //worksheet3r.Cell(5 + i, 7).Value = ukupp;
                    //worksheet3r.Cell(5 + i, 8).Value = ukupkol;
                    //worksheet3r.Cell(6 + i, 7).Value = "Postotak realizacije";

                    //worksheet3r.Cell(6 + i, 8).Value = Math.Round(((ukupkol + 0.001) / (ukupp + 0.001) * 100), 2);
                }
                else
                {// drugi ldp po radncima

                    if (1 == 1)
                    {
                        sql1 = "select * from rfind.dbo.evidnormiradad('" + datLDP + "','" + datLDP + "') order by vrsta,hala,smjena,linija"; //" + smj; bez računanja efektivnog radnog vremena  

                        // sql1 = "rfind.dbo.ldp_recalc '" + dat13 + "','" + dat23 + "',41";
                        SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                        SqlDataReader reader = sqlCommand.ExecuteReader();
                        string pec;
                        int i = 0;
                        int ukupn = 0, ukupp = 0, ukupkol = 0, zastoj = 0, norma = 0;
                        string vrsta, hala, smjena, linija, radnik;
                        string vrsta1 = "", hala1 = "", smjena1 = "", linija1 = "", radnik1 = "", radnikid1 = "";
                        int kolicina = 0, prazni = 0;
                        while (reader.Read())
                        {
                            prazni = 0;
                            radnik = reader["radnik"].ToString().Trim();
                            vrsta = (reader["vrsta"].ToString());
                            if (vrsta == "ND")
                                continue;


                            if (radnik.Length == 0)
                            {
                                int z = 0;
                            }
                            if (radnik == "")
                            {
                                vrsta = (reader["vrsta"].ToString());
                                hala = (reader["hala"].ToString());
                                smjena = (reader["smjena"].ToString());
                                kolicina = (int.Parse)(reader["smjena"].ToString());
                                linija = (reader["linija"].ToString());

                                if (vrsta == vrsta1 && hala == hala1 && smjena == smjena1 && linija == linija1 && kolicina > 0)
                                {
                                    radnik = radnik1;
                                    prazni = 1;
                                }
                                else
                                {
                                    continue;
                                }


                            }

                            //radnik = (reader["radnik"].ToString());                            
                            //vrsta = (reader["vrsta"].ToString());
                            //hala = (reader["hala"].ToString());
                            //smjena = (reader["smjena"].ToString());
                            //linija = (reader["linija"].ToString());


                            Console.WriteLine("Pregled po radnicima  trenutno vrijeme " + DateTime.Now);
                            string sk1 = reader["kolicinaok"].ToString();
                            if (reader["kolicinaok"] == DBNull.Value)
                            {
                                k1 = 0;
                            }
                            else
                            {
                                k1 = (int.Parse)(reader["kolicinaok"].ToString());
                            }

                            if (reader["norma"] == DBNull.Value)
                            {
                                n1 = 0;
                            }
                            else
                            {
                                n1 = (int.Parse)(reader["norma"].ToString());
                            }
                            double post1 = 0.0;
                            if (k1 > 0)
                            {
                                post1 = (k1 + 0.00001) / (n1 + 0.0001);
                            }
                            else
                            {
                                post1 = 0.0;
                            }

                            worksheet3r2.Cell(5 + i, 1).Value = radnik;
                            worksheet3r2.Cell(5 + i, 2).Value = (reader["vrsta"].ToString());
                            worksheet3r2.Cell(5 + i, 3).Value = (reader["hala"].ToString());
                            worksheet3r2.Cell(5 + i, 4).Value = (reader["smjena"].ToString());
                            worksheet3r2.Cell(5 + i, 5).Value = (reader["linija"].ToString());

                            worksheet3r2.Cell(5 + i, 6).Value = reader["nazivpar"];
                            worksheet3r2.Cell(5 + i, 7).Value = reader["brojrn"];
                            worksheet3r2.Cell(5 + i, 8).Value = (reader["proizvod"].ToString());
                            worksheet3r2.Cell(5 + i, 9).Value = (reader["norma"].ToString());

                            worksheet3r2.Cell(5 + i, 11).Value = (reader["kolicinaok"].ToString());
                            double satirada = (double.Parse)((reader["minutaradaradnika"].ToString()).Replace(',', '.')) / 100;
                            //string brisi = reader["trajanje"].ToString();

                            //if (brisi.Length > 0)
                            //    zastoj = (int.Parse)(reader["trajanje"].ToString());
                            //else
                            //    zastoj = 0;

                            string brisi = reader["norma"].ToString();
                            if (brisi.Length > 0)
                                norma = (int.Parse)(reader["norma"].ToString());
                            else
                                norma = 0;

                            // int srada = (int.Parse)(satirada);
                            worksheet3r2.Cell(5 + i, 13).Value = (reader["otpadobrada"].ToString());
                            worksheet3r2.Cell(5 + i, 14).Value = (reader["otpadmat"].ToString());
                            worksheet3r2.Cell(5 + i, 15).Value = reader["minutaradaradnika"];
                            worksheet3r2.Cell(5 + i, 16).Value = (reader["napomena1"].ToString());
                            worksheet3r2.Cell(5 + i, 17).Value = (reader["napomena2"].ToString());
                            worksheet3r2.Cell(5 + i, 18).Value = (reader["napomena3"].ToString());
                            //    double planirano = (satirada * 60.0 - zastoj) * norma / (satirada*60);

                            double planirano = norma * (satirada / 480);

                            //int n1 = (int.Parse)(reader["norma"].ToString());
                            double posto1 = 0.00;
                            n1 = (int)(planirano);

                            if (n1 > 0)
                            {
                                if (k1 == 0)
                                {
                                    posto1 = 0.00;
                                }
                                else
                                {
                                    posto1 = k1 / (n1 + 0.0000001);
                                }

                                if (posto1 >= 1.00)
                                {

                                    worksheet3r2.Cell(5 + i, 12).Style.Fill.BackgroundColor = XLColor.GreenPigment;
                                }
                                if (posto1 < 0.900)
                                {

                                    worksheet3r2.Cell(5 + i, 12).Style.Fill.BackgroundColor = XLColor.Red;
                                }
                                if (posto1 >= 0.900 && posto1 < 1.00)
                                {

                                    worksheet3r2.Cell(5 + i, 12).Style.Fill.BackgroundColor = XLColor.Orange;
                                }
                            }

                            worksheet3r2.Cell(5 + i, 10).Value = (int)planirano;
                            worksheet3r2.Cell(5 + i, 12).Value = posto1;

                            if (i % 2 == 0)
                                worksheet3r2.Range(8 + i, 1, 8 + i, 19).Style.Fill.BackgroundColor = XLColor.WhiteSmoke;
                            else
                                worksheet3r2.Range(8 + i, 1, 8 + i, 19).Style.Fill.BackgroundColor = XLColor.White;

                            ukupn = ukupn + n1;

                            if (reader["norma"] != DBNull.Value)
                            {
                                ukupp = ukupp + n1;
                                ukupkol = ukupkol + k1;
                            }

                            i++;
                            if (prazni == 0)
                            {
                                radnik1 = (reader["radnik"].ToString());
                            }
                            else
                            {
                                //radnik1 = "xxx";
                            }
                            vrsta1 = (reader["vrsta"].ToString());
                            hala1 = (reader["hala"].ToString());
                            smjena1 = (reader["smjena"].ToString());
                            linija1 = (reader["linija"].ToString());
                            

                        }
                        worksheet3r2.Cell(5 + i, 9).Value = "Ukupno :";
                        worksheet3r2.Cell(5 + i, 10).Value = ukupn;
                        worksheet3r2.Cell(5 + i, 11).Value = ukupkol;
                        worksheet3r2.Cell(6 + i, 9).Value = "Postotak realizacije";

                        worksheet3r2.Cell(6 + i, 10).Value = Math.Round(((ukupkol + 0.001) / (ukupp + 0.001) * 100), 2);
                    }
                }
                cn.Close();
            }

            Console.WriteLine("Pregled po radnicima završen trenutno vrijeme " + DateTime.Now);
            smj = "5";  // po linijama

            using (SqlConnection cn = new SqlConnection(connectionString))
            {
                cn.Open();
                //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                sql1 = "rfind.dbo.ldp_recalc '" + datLDP + "','" + datLDP + "'," + smj; ;
                SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                SqlDataReader reader = sqlCommand.ExecuteReader();
                string pec;
                int i = 0;
                int ukupn = 0, ukupp = 0, ukupkol = 0;
                string hala1 = "", smjena1 = "", proiz1 = "", linija1 = "", radninalog1 = "", radninalog2 = "";
                string hala2 = "", smjena2 = "", proiz2 = "", linija2 = "";
                string kupac1 = "", kupac2 = "", naziv1 = "" ;
                int norma1 = 0, kolicina1 = 0, kolicina2 = 0, norma2 = 0, norma = 0, kolicina = 0;
                int prvii = 1, normaukup1 = 0, normaukup2 = 0, erv1 = 0, erv2 = 0, erv = 0, normaukupp = 0;
                double cijena1 = 0.0, cijena2 = 0.0, norm = 0.0, norm1 = 0.0;
                bool joss = true;


                if (reader.Read())
                {
                    var loop = true;
                    while (loop)
                    {

                        if (i % 2 == 0)
                            worksheet4l.Range(8 + i, 1, 8 + i, 15).Style.Fill.BackgroundColor = XLColor.WhiteSmoke;
                        else
                            worksheet4l.Range(8 + i, 1, 8 + i, 15).Style.Fill.BackgroundColor = XLColor.White;


                        hala1 = (reader["hala"].ToString());
                        if ( hala1=="" )
                        {
                            hala1 = (reader["halal"].ToString());
                            naziv1= (reader["naziv"].ToString());
                            worksheet4l.Cell(5 + i, 1).Value = dat1;
                            worksheet4l.Cell(5 + i, 2).Value = hala1;
                            worksheet4l.Cell(5 + i, 3).Value = naziv1;

                            i++;
                            reader.Read();
                            continue;
                        }

                        if (i % 2 == 0)
                            worksheet4l.Range(8 + i, 1, 8 + i, 15).Style.Fill.BackgroundColor = XLColor.WhiteSmoke;
                        else
                            worksheet4l.Range(8 + i, 1, 8 + i, 15).Style.Fill.BackgroundColor = XLColor.White;

                        smjena1 = (reader["smjena"].ToString());
                        proiz1 = (reader["nazivpro"].ToString());
                        linija1 = (reader["linija"].ToString());
                        radninalog1 = (reader["brojrn"].ToString());
                        kolicina1 = (int.Parse)(reader["kolicinaok"].ToString());
                        normaukup1 = (int.Parse)(reader["normukup"].ToString());
                        erv1 = (int.Parse)(reader["erv"].ToString());
                        norma1 = (int.Parse)(reader["norma"].ToString());
                        cijena1 = (double.Parse)(reader["cijena"].ToString());
                        kupac1 = (reader["kupac"].ToString());
                        norm1 = norma1 * erv1 / normaukup1;
                        worksheet4l.Cell(5 + i, 1).Value = reader["datum"];
                        string brisi1 = (reader["erv"].ToString());

                        loop = reader.Read();

                        if ((hala1 == hala2 && linija1 == linija2 && proiz1 == proiz2) || (prvii == 1))
                        {
                            norma    = norma    + norma1    ;
                            kolicina = kolicina + kolicina1 ;
                            erv      = erv      + erv1      ;
                            normaukupp = normaukupp + normaukup1;
                            norm       = norm + norma1 * erv1 / normaukup1;

                            if (hala1 == hala2 && smjena1 == smjena2 && proiz1 == proiz2 && linija1 == linija2 && radninalog1 != radninalog2)
                            {
                                norma = norma - norma1;
                                // norm  = norm  - norma1 * erv1 / normaukup1;
                                // erv   = erv   - erv1;
                                normaukupp = normaukupp - normaukup1;
                            }
                            norm       = norma * erv / normaukupp;
                            norma1     = norma;
                            erv1       = erv;
                            normaukup1 = normaukupp;
                            norm1      = norm;

                            hala2 = hala1;
                            erv2  = erv1;

                            smjena2     = smjena1;
                            proiz2      = proiz1;
                            radninalog2 = radninalog1;
                            linija2     = linija1;
                            cijena2     = cijena1;
                            kupac2      = kupac1;
                            prvii       = 0;

                            if (!loop)
                            {
                                loop = loop;
                                //You are on the last record. Use values read in 1.
                                //Do some exceptions
                            }
                            else
                            {
                                continue;
                                //You are not on the last record.
                                //Process values read in 1., e.g. myvar
                            }

                        }

                        worksheet4l.Cell(5 + i, 2).Value = hala2;
                        worksheet4l.Cell(5 + i, 3).Value = linija2;
                        if (linija2.Contains("VC510") || linija2.Contains("HURCO") || linija2.Contains("GT600"))
                        {
                            worksheet4l.Cell(5 + i, 4).Value = "Dodatne operacije";
                        }
                        else
                        {
                            worksheet4l.Cell(5 + i, 4).Value = "Tokarenje";
                        }

                        worksheet4l.Cell(5 + i, 5).Value = proiz2;
                        worksheet4l.Cell(5 + i, 6).Value = norma; //  reader["norma"]; ;  // norma
                        worksheet4l.Cell(5 + i, 7).Value = kolicina; // reader["kolicinaok"]; ;  // komada                    

                        //    norma1 = norma;

                        erv = 0;

                        if (brisi1.Length > 0)
                        {
                            erv = (int)((double.Parse)(brisi1));
                        }
                        else
                        {
                            erv = 0;
                        }
                        double ostvareno = 0.0; double planirano = 0.0, normaukup = 0.0;

                        //if (norma1 != 0)
                        {
                            ostvareno = kolicina * cijena2;
                            planirano = norma * cijena2;  //  za normu od 480 minuta
                                                          //normaukup = normaukup2; // (double.Parse)(reader["normukup"].ToString());
                        }

                        double posto1 = 0.00;
                        double n1 = 0.0, k1 = 0.0;
                        k1 = ostvareno;
                        n1 = (planirano) * erv / normaukupp;    // korigirina norma, normaukup = ukupno normirano vrijeme
                        n1 = norm * cijena2;

                        worksheet4l.Cell(5 + i, 8).Value = norm; // planirano komada
                        worksheet4l.Cell(5 + i, 9).Value = cijena2;

                        posto1 = (k1 + 0.0000001) / (n1 + 0.0000001);

                        if (k1 <= 0.0001 || n1 <= 0.0001)
                        {
                            posto1 = 0.0;
                        }

                        if (posto1 >= 1.00)
                        {
                            worksheet4l.Cell(5 + i, 12).Style.Fill.BackgroundColor = XLColor.Green;
                        }
                        if (posto1 < 0.900)
                        {
                            worksheet4l.Cell(5 + i, 12).Style.Fill.BackgroundColor = XLColor.Red;
                        }
                        if (posto1 >= 0.900 && posto1 < 1.00)
                        {
                            worksheet4l.Cell(5 + i, 12).Style.Fill.BackgroundColor = XLColor.Orange;
                        }
                        
                        worksheet4l.Cell(5 + i, 10).Value = ostvareno;
                        worksheet4l.Cell(5 + i, 11).Value = n1;   // korigirana norma za srtvarno vrijeme rada
                        worksheet4l.Cell(5 + i, 12).Value = posto1;
                        worksheet4l.Cell(5 + i, 13).Value = ostvareno - planirano;
                        //   worksheet4l.Cell(5 + i, 12).Value = (reader["napomena"].ToString());
                        worksheet4l.Cell(5 + i, 15).Value = kupac2;

                        i++;
                        hala2 = hala1;
                        smjena2 = smjena1;
                        proiz2 = proiz1;
                        radninalog2 = radninalog1;
                        linija2 = linija1;
                        cijena2 = cijena1;
                        norma = norma1;
                        norm = norm1;
                        erv = erv1;
                        normaukupp = normaukup1;
                        kolicina = kolicina1;
                        kupac2 = kupac1;
                        prvii = 0;
                        if (i % 2 == 0)
                            worksheet4l.Range(8 + i, 1, 8 + i, 15).Style.Fill.BackgroundColor = XLColor.WhiteSmoke;
                        else
                            worksheet4l.Range(8 + i, 1, 8 + i, 15).Style.Fill.BackgroundColor = XLColor.White;
                    }
                }
                cn.Close();
            }
            Console.WriteLine("Pregled po linijama  trenutno vrijeme " + DateTime.Now);

            smj = "6";  //   Štelanje
            dat13 = datLDP + " 06:00:00";
            dat23 = dat3 + " 06:00:00";

            using (SqlConnection cn = new SqlConnection(connectionString))
            {
                cn.Open();
                //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                sql1 = "rfind.dbo.ldp_recalc '" + dat13 + "','" + dat23 + "'," + smj; ;
                SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                SqlDataReader reader = sqlCommand.ExecuteReader();
                string pec;
                int i = 0;
                int ukupn = 0, ukupp = 0, ukupkol = 0, erv = 0;
                workshstelanje.Cell(4, 3).Value = workshstelanje.Cell(4, 3).Value + "  " + datreps;

                while (reader.Read())
                {
                    workshstelanje.Cell(9 + i, 1).Value = (reader["hala"].ToString());
                    workshstelanje.Cell(9 + i, 2).Value = (reader["naziv"].ToString());
                    workshstelanje.Cell(9 + i, 3).Value = (reader["aktivnost"].ToString());
                    workshstelanje.Cell(9 + i, 4).Value = reader["pocetak"]; ;  // norma
                    workshstelanje.Cell(9 + i, 5).Value = reader["kraj"]; ;  // komada                    
                    workshstelanje.Cell(9 + i, 6).Value = reader["trajanje_minuta"];
                    workshstelanje.Cell(9 + i, 7).Value = reader["izgubljeno"];
                    workshstelanje.Cell(9 + i, 8).Value = reader["brojrn"];
                    workshstelanje.Cell(9 + i, 9).Value = reader["napomena"];
                    workshstelanje.Cell(9 + i, 10).Value = reader["username"];
                    workshstelanje.Cell(9 + i, 11).Value = reader["vrstanarudzbe"];
                    workshstelanje.Cell(9 + i, 12).Value = reader["nazivpar"];
                    if (i % 2 == 0)
                        workshstelanje.Range(8 + i, 1, 8 + i, 12).Style.Fill.BackgroundColor = XLColor.WhiteSmoke;
                    else
                        workshstelanje.Range(8 + i, 1, 8 + i, 12).Style.Fill.BackgroundColor = XLColor.White;

                    i++;

                }

                cn.Close();
            }
            Console.WriteLine("Pregled izostanaka  trenutno vrijeme " + DateTime.Now);

            //workshstelanje.Cell(29, 3).Value



            //            string fileName = @"C:\brisi\dsr20062017.xlsm";

            var fi = new FileInfo(fileName);
            if (fi.Exists) File.Delete(fileName);

            workbook.SaveAs(fileName);

            //file bez vrijednosti
            string fileNamebv = @"c:\brisi\dsrv" + nuland + d2.Day.ToString() + nulanm + d2.Month.ToString() + d2.Year.ToString() + ".xlsx";
            fi = new FileInfo(fileNamebv);
            if (fi.Exists) File.Delete(fileNamebv);

            // ovja puta brisi sve vrijednosti u EUR
            var worksheetbv = workbook.Worksheet(1);
            worksheetbv.Cell(29, 3).Value = "";    // ukupna vrijednost
            // kaliona
            worksheetbv.Cell(33, 11).Value = "";    // realizacija
            worksheetbv.Cell(34, 11).Value = "";    // realizacija

            //Range startCell = worksheetbv.Cell(43, 2];  // vrijednost repromaterijala
            //Range endCell = worksheetbv.Cell(43, 16];
            // With a string

            //var range1 = worksheetbv.Range(43,2,43,16);
            //range1.Cell(1, 1).Value = "ws.Range(\"B43:P43\").Merge()";
            //range1.Merge();
            //range1.Value = "";
            //worksheetbv.Range[startCell, endCell).Value = "";

            worksheetbv.Range("B43:P43").Clear();
            //startCell = worksheetbv.Cell(7, 6];  // vrijednost obrade u eur
            //endCell = worksheetbv.Cell(22, 7];

            //worksheetbv.Range[startCell, endCell).Value = "";

            //range1 = worksheetbv.Range("F7:G22");
            //range1.Cell(1, 1).Value = "ws.Range(\"F7:G22\").Merge()";
            //range1.Merge();
            worksheetbv.Range("F7:G22").Clear();

            var worksheetbvt = workbook.Worksheet(2);  // tjedni izvještaj
            worksheetbvt.Cell(29, 3).Value = "";    // ukupna vrijednost u eur
            worksheetbvt.Cell(33, 9).Value = "";    // kaliona realizacija

            var worksheetbvlinija = workbook.Worksheet(4);  // po linijama
            //startCell = worksheetbvlinija.Cell(5, 9];  // kolona sa cijenama
            //endCell = worksheetbvlinija.Cell(290, 9];

            //worksheetbvlinija.Range[startCell, endCell).Value = "";

            //range1 = worksheetbvlinija.Range("I9:K250");
            //range1.Cell(1, 1).Value = "ws.Range(\"I9:K250\").Merge()";
            //range1.Merge();
            worksheetbv.Range("I9:K250").Clear();
            worksheetbv.Range("M9:M250").Clear();

            //range1 = worksheetbvlinija.Range("M9:M250");
            //range1.Cell(1, 1).Value = "ws.Range(\"M9:M250\").Merge()";
            //range1.Merge();

            //startCell = worksheetbvlinija.Cell(5, 10];  // kolona sa ostvarenim eurima
            //endCell = worksheetbvlinija.Cell(290, 10];

            //worksheetbvlinija.Range[startCell, endCell).Value = "";

            //startCell = worksheetbvlinija.Cell(5, 11];  // kolona sa planiranim eurima
            //endCell = worksheetbvlinija.Cell(290, 11];

            //worksheetbvlinija.Range[startCell, endCell).Value = "";


            //startCell = worksheetbvlinija.Cell(5, 13];  // kolona sa  razlikom u eurima
            //endCell = worksheetbvlinija.Cell(290, 13];

            //worksheetbvlinija.Range[startCell, endCell).Value = "";

            workbook.SaveAs(fileNamebv);


            Console.WriteLine("Snimljen file ddmmyyyy   trenutno vrijeme " + DateTime.Now);
            //Console.ReadKey();
            //workbook.Close(false);
            //excel.Application.Quit();
            //excel.Quit()
            //workbook.Close(true, Type.Missing, Type.Missing);
                      
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            workshstelanje = null;
            workbook = null;
            

//MailMessage mail = new MailMessage("gasparic.s@feroimpex.hr", "legac.b@feroimpex.hr,legac.h@feroimpex.hr,legac.z@feroimpex.hr,horvatic.v@feroimpex.hr,kavalin.k@feroimpex.hr,gasparic.s@feroimpex.hr");            
MailMessage mail = new MailMessage("gasparic.s@feroimpex.hr", "gasparic.s@feroimpex.hr,srecckog@gmail.com");

            SmtpClient client = new SmtpClient();
            client.Port = 25;
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            client.UseDefaultCredentials = false;
            client.Host = "smtp.metronet.hr";
            mail.Subject = "DPR " + datreps;
            mail.Body = "Daily production report for " + datreps;

            Attachment attachment = new Attachment(fileName, System.Net.Mime.MediaTypeNames.Application.Octet);
            System.Net.Mime.ContentDisposition disposition = attachment.ContentDisposition;
            disposition.CreationDate = File.GetCreationTime(fileName);
            disposition.ModificationDate = File.GetLastWriteTime(fileName);
            disposition.ReadDate = File.GetLastAccessTime(fileName);
            disposition.FileName = Path.GetFileName(fileName);
            disposition.Size = new FileInfo(fileName).Length;
            disposition.DispositionType = System.Net.Mime.DispositionTypeNames.Attachment;

            mail.Attachments.Add(attachment);
            client.Send(mail);
            client.Dispose();

            // ovaj puta šalji bez vrijednosti u EUR
//mail = new MailMessage("gasparic.s@feroimpex.hr", "sestan.p@feroimpex.hr,kicin.d@feroimpex.hr,grgecic.d@feroimpex.hr,gasparic.s@feroimpex.hr,srecckog@gmail.com,stefanac.d@feroimpex.hr,darapi.i@feroimpex.hr,janjic.d@feroimpex.hr,janjcic.d@feroimpex.hr, scuric.i@feroimpex.hr,biscan.s@feroimpex.hr");            
 mail = new MailMessage("gasparic.s@feroimpex.hr", "gasparic.s@feroimpex.hr,srecckog@gmail.com");

            client = new SmtpClient();
            client.Port = 25;
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            client.UseDefaultCredentials = false;
            client.Host = "smtp.metronet.hr";

            mail.Subject = "_DPR " + datreps;
            mail.Body = "_Daily production report for " + datreps;


            attachment = new Attachment(fileNamebv, System.Net.Mime.MediaTypeNames.Application.Octet);
            System.Net.Mime.ContentDisposition disposition2 = attachment.ContentDisposition;
            disposition2.CreationDate = File.GetCreationTime(fileNamebv);
            disposition2.ModificationDate = File.GetLastWriteTime(fileNamebv);
            disposition2.ReadDate = File.GetLastAccessTime(fileNamebv);
            disposition2.FileName = Path.GetFileName(fileNamebv);
            disposition2.Size = new FileInfo(fileNamebv).Length;
            disposition2.DispositionType = System.Net.Mime.DispositionTypeNames.Attachment;

            mail.Attachments.Add(attachment);
            client.Send(mail);
            client.Dispose();

            var processes = from p in System.Diagnostics.Process.GetProcessesByName("EXCEL")
                            select p;

            foreach (var process in processes)
            {
                int z = 0;
                if (process.MainWindowTitle.Contains("Microsoft Excel"))
                    process.Kill();
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();

        }
        //static int korekcijalinija( string dat1,string dat2)
        //{
        //    using (SqlConnection cn = new SqlConnection(connectionString))
        //    {
        //        cn.Open();
        //        //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
        //        // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
        //        string sql1 = "select hala, linija, count(*) kol from(select distinct kupac, hala, linija, VrstaNarudzbe from feroapp.dbo.evidnormi('"+dat1+"', "+dat2+", 0) where gotovo = 1 and obrada3 = 0 and KolicinaOK >= 0 ) x1 group by hala, linija having count(*)> 1";


        //        SqlCommand sqlCommand = new SqlCommand(sql1, cn);

        //        SqlDataReader reader = sqlCommand.ExecuteReader();
        //        string kupac, vnar = "";

        //        while (reader.Read())
        //        {
        //            kupac = reader["Hala"].ToString();
        //            vnar = reader["linija"].ToString();




        //        }

        //         //   if (kupac.Contains("Austria"))


        //                return 1;
        //}

    }

}
