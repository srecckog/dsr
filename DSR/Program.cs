using System;
using System.Globalization;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
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

            string connectionString = @"Data Source=192.168.0.3;Initial Catalog=FeroApp;User ID=sa;Password=AdminFX9.";
            string connectionString2 = @"Data Source=192.168.0.3;Initial Catalog=RFIND;User ID=sa;Password=AdminFX9.";

            string c7 = "0", c8 = "0", c9 = "0", c10 = "0", c11 = "0", c12 = "0", c13 = "0", c14 = "0", c15 = "0",  c151 = "0", c16 = "0", c17 = "0" , c18 = "0" , c19 = "0", c21 = "0" ,c23="0"; // 
            double f7 = 0.0, f8 = 0.0, f9 = 0.0, f10 = 0.0, f11 = 0.0, f12 = 0.0, f13 = 0.0, f14 = 0.0, f15 = 0.0, f151 = 0.0, f16 = 0.0, f17 = 0.0, f18 = 0.0, f19 = 0.0, f21 = 0.0, f23 = 0.0; // 
            double g7 = 0.0, g8 = 0.0, g9 = 0.0, g10 = 0.0, g11 = 0.0, g12 = 0.0, g13 = 0.0, g14 = 0.0, g15 = 0.0,  g151 = 0.0, g16 = 0.0, g17 = 0.0,  g18=0.0 , g19 = 0.0, g21 = 0.0, g23 = 0.0 ; // 

            double i7 = 0.0, i8 = 0.0, i9 = 0.0, i10 = 0.0, i11 = 0.0, i12 = 0.0, i13 = 0.0, i14 = 0.0, i15 = 0.0,  i151 = 0.0, i16 = 0.0, i17 = 0   , i18=0   , i19 = 0.0, i21 = 0.0 , i22=0.0 ,i23=0.0; // 
            string j7 = "0", j8 = "0", j9 = "0", j10 = "0", j11 = "0", j12 = "0", j13 = "0", j14 = "0", j15 = "0",  j151 = "0", j16 = "0", j17="0"   , j18 ="0",j19 = "0", j21 = "0" , j23="0"; // 
            double    k7 = 0.0, k8 = 0.0, k9 = 0.0, k10 = 0.0, k11 = 0.0, k12 = 0.0, k13 = 0.0, k14 = 0.0, k15 = 0.0,  k151 = 0.0, k16 = 0.0, k17=0.0   , k18 = 0.0, k19 = 0.0, k21 = 0.0 , k23=0.0; // 
            double    q7 = 0.0, q8 = 0.0 ,q9 = 0.0, q10 = 0.0, q11 = 0.0, q12 = 0.0, q13 = 0.0, q14 = 0.0, q15 = 0.0,  q151 = 0.0, q16 = 0.0, q17=0.0   , q18 = 0.0, q19 = 0.0, q21 = 0.0 , q23=0.0 ; // 
            string sql1 = "";
            string dat1 = "2017-06-20", dat2 = "2017-06-20", dat3 = "";
            string vo = "";
            string kupac7 = "274", kupac8 = "121269", kupac9 = "221452", kupac10 = "221453", kupac11 = "121274", kupac12 = "121274", kupac13 = "121273", kupac14 = "8752", kupac15 = "121301", kupac151 = "121302", kupac16 = "45150", kupac17 = "121273", kupac18 = "221463", kupac19 = "121278", kupac21 = "121400",kupac23="0";
            string nkupac7 = "BDF -274", nkupac8 = "BMW - 121269", nkupac9 = "DEBRECEN - 221452", nkupac10 = "Brašov - 221453", nkupac11 = "Wupertal prsten - 121274", nkupac12 = "Wupertal valjci - 121274", nkupac13 = "Kysuce kučišta - 121273", nkupac14 = "SKF 8752", nkupac15 = "Sona - 121301", nkupac151 = "Sona - 121302" ,nkupac16 = "Sigma prsteni - 45150", nkupac17 = "Kysuce prsteni - 121273", nkupac18 = "JLG - 221463", nkupac19 = "Brasil - 121278", nkupac21 = "NSK - 121400" , nkupac23 ="";
            int IndexCNC = 0;

            DateTime d1 = new DateTime(2017, 6, 20);
            DateTime d10 = new DateTime(2017, 7, 27);
            DateTime d2 = DateTime.Now.AddDays(-1);
            //DateTime jucerd = d2;    // jučerašnji dan
            //DateTime d2 = DateTime.Now;
            if (d2.Hour < 12)
            {
//                           d2 = d2.AddDays(-1);
            }
            DateTime d3 = DateTime.Now;

            if (1 == 2)
            {
                d2 = new DateTime(2019, 7, 12 ) ;
                d3 = new DateTime(2019, 7, 12 ) ;
            }
            DateTime jucerd = d2;    // jučerašnji dan
            string dan1 = d2.Day.ToString();
            string m1 = d2.Month.ToString();
            string g1 = d2.Year.ToString();

            dat1 = d2.Year.ToString() + '-' + d2.Month.ToString() + '-' + d2.Day.ToString();

            string datjucer = dat1;            
            string datdanas = d2.AddDays(1).Year.ToString()+ '-' + d2.AddDays(1).Month.ToString() + '-' + d2.AddDays(1).Day.ToString();

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

            Application excel = new Application();
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            //Workbook workbook = excel.Workbooks.Open(@"C:\brisi\dsr_template33.xlsx", ReadOnly: false, Editable: true);
            //Workbook workbook = excel.Workbooks.Open(@"C:\brisi\dsr_template1001.xlsx", ReadOnly: false, Editable: true);

            //    Workbook workbook = excel.Workbooks.Open(@"C:\brisi\dsr_template2203.xlsx", ReadOnly: false, Editable: true);

            //Workbook workbook = excel.Workbooks.Open(@"C:\brisi\dsr_template1907.xlsx", ReadOnly: false, Editable: true);
            //Workbook workbook = excel.Workbooks.Open(@"C:\brisi\dsr_template2007.xlsx", ReadOnly: false, Editable: true);
            Workbook workbook = excel.Workbooks.Open(@"C:\brisi\dsr_template1803.xlsx", ReadOnly: false, Editable: true);

            //            Worksheet works10 = workbook.Worksheets.Item[10] as Worksheet;   // ldp po radnicima

            // for (int ws = 1; ws <= 5; ws++)
            for (int ws = 1; ws <= 5; ws++)
                {

                if (ws == 2 || ws==3)
                {
                    if (ws == 2)   // tjedni
                    {
                        dat10 = monday.Day.ToString() + "." + monday.Month.ToString() + "." + monday.Year.ToString() + " - " + sunday.Day.ToString() + "." + sunday.Month.ToString() + "." + sunday.Year.ToString();
                        dat1 = monday.Year.ToString() + "-" + monday.Month.ToString() + "-" + monday.Day.ToString();
                        dat2 = sunday.Year.ToString() + "-" + sunday.Month.ToString() + "-" + sunday.Day.ToString();
                    }
                    else    // mjesečni
                    {
                        dat10 = "01." + jucerd.Month.ToString() + "." +jucerd.Year.ToString() + " - " + jucerd.Day.ToString() + "." + jucerd.Month.ToString() + "." + jucerd.Year.ToString();
                        dat1 = jucerd.Year.ToString() + "-" + jucerd.Month.ToString() + "-01" ;
                        dat2 = jucerd.Year.ToString() + "-" + jucerd.Month.ToString() + "-" + jucerd.Day.ToString();
                    }
                }
                else  // dnevni izvještaj
                {
                    dat10 = dat100;
                    d2 = jucerd;
                    dat1 = d2.Year.ToString() + '-' + d2.Month.ToString() + '-' + d2.Day.ToString();
                    dat2 = dat1;                    
                }

                string mm1 = "";   // dodaj 0 na pocetak mjeseca
                if (d2.Month <= 9)
                {
                    mm1 = "0";
                }
                string mmgggg1 = mm1 + d2.Month.ToString() + '-' + d2.Year.ToString();
                

                if (1 == 1)
                {

                    using (SqlConnection cn = new SqlConnection(connectionString))
                    {
                        cn.Open();
                        //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                        // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);

                        sql1 = "rfind.dbo.realizacija '" + dat1 + "','" + dat2 + "',21";

                        SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                        SqlDataReader reader = sqlCommand.ExecuteReader();
                        string kupac, idpar;
                        if (reader.HasRows)
                        {
                            IndexCNC = 1;
                        }
                                                                       
                    }
                                                                          
                            // ostvareno komada
                            using (SqlConnection cn = new SqlConnection(connectionString))
                    {
                        cn.Open();
                        //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                        // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);

                        sql1 = "rfind.dbo.realizacija '" + dat1 + "','" + dat2 + "',1";

                        SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                        SqlDataReader reader = sqlCommand.ExecuteReader();
                        string kupac,idpar;

                        while (reader.Read())
                        {
                            kupac = reader["Kupac"].ToString();
                            idpar = reader["id_par"].ToString();
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

                                if (kupac.Contains("Valj"))  // wupertal
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

                            if (kupac.Contains("SONA BLW"))
                            {
                                if (idpar == "121301")
                                {
                                    c15 = (reader["Kolicinaok"].ToString());
                                    //j15 = (reader["broj_stelanja"].ToString());  // 
                                    ukupv = ukupv + (double.Parse)(reader["vrijednost"].ToString());
                                    ukupk = ukupk + (double.Parse)(reader["kolicinaok"].ToString());
                                }
                                if (idpar == "121302")
                                {
                                    c151 = (reader["Kolicinaok"].ToString());
                                    //j15 = (reader["broj_stelanja"].ToString());  // 
                                    ukupv = ukupv + (double.Parse)(reader["vrijednost"].ToString());
                                    ukupk = ukupk + (double.Parse)(reader["kolicinaok"].ToString());
                                }

                            }


                            if (kupac.Contains("SIGMA"))  // wupertal
                            {
                                c16 = (reader["Kolicinaok"].ToString());
                                //j16 = (reader["broj_stelanja"].ToString());  // 
                                ukupv = ukupv + (double.Parse)(reader["vrijednost"].ToString());
                                ukupk = ukupk + (double.Parse)(reader["kolicinaok"].ToString());
                            }

                            if (kupac.Contains("JLG"))  // JLG
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

                            if (kupac.Contains("xxx"))  // 
                            {
                                c23 = (reader["Kolicinaok"].ToString());
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
                j7 = "0"; j8 = "0"; j9 = "0"; j10 = "0"; j11 = "0"; j12 = "0"; j13 = "0"; j14 = "0"; j15 = "0"; j16 = "0"; j17 = "0" ; j19 = "0"; j21 = "0"; j23="0"; // 
                using (SqlConnection cn = new SqlConnection(connectionString))
                {
                    cn.Open();
                    //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id;[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                    // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                    if (ws <= 3)
                    {
                        sql1 = "rfind.dbo.LDP_recalc '" + dat1 + "','" + dat3 + "',102";
                    }
                    else
                    {
                        sql1 = "rfind.dbo.LDP_recalc '" + dat1 + "','" + dat3 + "',1020";
                    }

                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string kupac, idpar,vrstanarudzbe,hala1="";

                    while (reader.Read())
                    {
                        if (ws >= 4)
                        {
                            hala1 = reader["hala"].ToString();
                            if ((ws == 4) && hala1.Contains("3")) // pušta samo halu 1
                            {
                                continue;
                            }
                            if ((ws == 5) && hala1.Contains("1"))  // pušta samo 3
                            {
                                continue;
                            }
                        }


                        kupac = reader["Kupac"].ToString();
                        idpar = reader["id_par"].ToString();
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

                            if (vrstanarudzbe.Contains("Valj"))  // wupertal
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
                            if (idpar == "121301")
                            {
                                j15 = (reader["broj_stelanja"].ToString());  // 
                            }
                            else
                            {
                                j151 = (reader["broj_stelanja"].ToString());  // 
                            }
                        }


                        if (kupac.Contains("SIGMA"))  // wupertal
                        {

                            j16 = (reader["broj_stelanja"].ToString());  // 

                        }

                        if (kupac.Contains("JLG"))  // JLG
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

                        if (kupac.Contains("xxx"))
                        {

                            j23 = (reader["broj_stelanja"].ToString());  // JLG

                        }


                    }
                    cn.Close();

                }
                Console.WriteLine("Izračunat broj štelanja  trenutno vrijeme " + DateTime.Now);


                // kupac, komada iz sELECT * FROM EvidencijaProizvedenoFakturirano_Zbirno('2017-08-02', '2017-08-02')
                c7 = "0"; c8 = "0"; c9 = "0"; c10 = "0"; c11 = "0"; c12 = "0"; c13 = "0"; c14 = "0"; c16 = "0" ; c17 = "0";c18 = "0"; c19 = "0"; c21 = "0"; c23 ="0"; // 
                f7 = 0.0; f8 = 0.0; f9 = 0.0; f10 = 0.0; f11 = 0.0; f12 = 0.0; f13 = 0.0; f14 = 0.0; f15 = 0.0 ; f151 = 0.0 ; f16 = 0.0; f17 = 0.0; f18 = 0.0;f19 = 0.0; f21 = 0.0; f23=0.0 ; // 
                g7 = 0.0; g8 = 0.0; g9 = 0.0; g10 = 0.0; g11 = 0.0; g12 = 0.0; g13 = 0.0; g14 = 0.0; g15 = 0   ; g151 = 0.0  ; g16 = 0.0; g17 = 0.0; g18 = 0.0; g19 = 0.0; g21 = 0.0 ; g23=0.0 ; // 

                string e7 = "0", e8 = "0", e9 = "0", e10 = "0", e11 = "0", e12 = "0", e13 = "0", e14 = "0", e15 = "0", e151="0", e16 = "0", e17 = "0", e18 = "0", e19 = "0", e21 = "0" , e23 ="0"; // 

                ukupk = int.Parse(c15);    // količina od sone je ok, nju zadržavam
                ukupk = 0;

                using (SqlConnection cn = new SqlConnection(connectionString))
                {
                    cn.Open();
                    //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as iddoubl;[ime x] as ime;[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                    // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);

                    // SELECT * FROM feroapp.dbo.EvidencijaProizvedenoFakturirano_Zbirno(@dat1, @dat2)
                    if (ws <=3)   // dnevni,tjedni,mjesečni
                    {
                        sql1 = "rfind.dbo.ldp_recalc '" + dat1 + "','" + dat2 + "',112";
                    }
                    else
                    {
                        sql1 = "rfind.dbo.ldp_recalc '" + dat1 + "','" + dat2 + "',113";
                    }

                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string kupac, idpar,vrstap,hala1="";

                    while (reader.Read())
                    {
                        if (ws >= 4)
                        {
                            hala1 = reader["hala"].ToString();
                            if ((ws == 4) && hala1.Contains("3")) // pušta samo halu 1
                            {
                                continue;
                            }
                            if ((ws == 5) && hala1.Contains("1"))  // pušta samo 3
                            {
                                continue;
                            }
                        }

                        kupac = reader["Kupac"].ToString();
                        idpar = reader["id_par"].ToString();
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

                            if (vrstap.Contains("Valj"))  // wupertal
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

                            if (idpar == "121301")
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
                            }
                            else
                            {
                                if (vo == "TOK")
                                {
                                    c151 = (reader["kolicina"].ToString());
                                    f151 = (double.Parse)(reader["vrijednost"].ToString());
                                }
                                else
                                {
                                    e151 = (reader["kolicina"].ToString());
                                    g151 = (double.Parse)(reader["vrijednost"].ToString());
                                }

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
                                f16 = (double.Parse)(reader["vrijednost"].ToString());

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

                        if (kupac.Contains("JLG"))  // JLG
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

                        if (kupac.Contains("xxx"))  // 


                        {
                            if (vo == "TOK")
                            {
                                c23 = (reader["kolicina"].ToString());
                                f23 = (double.Parse)(reader["vrijednost"].ToString());

                            }
                            else
                            {
                                e23 = (reader["kolicina"].ToString());
                                g23 = (double.Parse)(reader["vrijednost"].ToString());

                            }

                            //e7 = (reader["Proizvedeno_do"].ToString());
                            ukupk = ukupk + (double.Parse)(reader["kolicina"].ToString());

                        }

                    }
                }

                // vrijednost
                if (1 == 2)
                {

                    using (SqlConnection cn = new SqlConnection(connectionString))
                    {
                        cn.Open();
                        //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as iddoubl;[ime x] as ime;[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                        // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);

                        // SELECT * FROM feroapp.dbo.EvidencijaProizvedenoFakturirano_Zbirno(@dat1, @dat2)
                        if (ws <= 3)   // dnevni,tjedni,mjesečni
                        {
                            sql1 = "rfind.dbo.ldp_recalc '" + dat1 + "','" + dat2 + "',1121";
                        }
                        else
                        {
                            sql1 = "rfind.dbo.ldp_recalc '" + dat1 + "','" + dat2 + "',1131";
                        }

                        SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                        SqlDataReader reader = sqlCommand.ExecuteReader();
                        string kupac, vrstap, hala1 = "";


                        while (reader.Read())
                        {
                            if (ws >= 4)
                            {
                                hala1 = reader["hala"].ToString();
                                if ((ws == 4) && hala1.Contains("3")) // pušta samo halu 1
                                {
                                    continue;
                                }
                                if ((ws == 5) && hala1.Contains("1"))  // pušta samo 3
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
                                    f7 = (double.Parse)(reader["vrijednost"].ToString());
                                }
                                else
                                {
                                    g7 = (double.Parse)(reader["vrijednost"].ToString());
                                }

                                //e7 = (reader["Proizvedeno_do"].ToString());

                            }

                            if (kupac.Contains("SCHWEINFURT"))
                            {
                                if (vo == "TOK")
                                {

                                    f8 = (double.Parse)(reader["vrijednost"].ToString());
                                }
                                else
                                {

                                    g8 = (double.Parse)(reader["vrijednost"].ToString());
                                }

                                //e7 = (reader["Proizvedeno_do"].ToString());

                            }

                            if (kupac.Contains("FAG"))
                            {
                                if (vo == "TOK")
                                {

                                    f9 = (double.Parse)(reader["vrijednost"].ToString());
                                }
                                else
                                {

                                    g9 = (double.Parse)(reader["vrijednost"].ToString());
                                }

                                //e7 = (reader["Proizvedeno_do"].ToString());

                            }

                            if (kupac.Contains("ROMANIA"))
                            {
                                if (vo == "TOK")
                                {

                                    f10 = (double.Parse)(reader["vrijednost"].ToString());
                                }
                                else
                                {

                                    g10 = (double.Parse)(reader["vrijednost"].ToString());
                                }

                                //e7 = (reader["Proizvedeno_do"].ToString());

                            }

                            if (kupac.Contains("Technologies"))  // wupertal
                            {
                                if (vrstap.Contains("Prsten"))  // wupertal
                                {
                                    if (vo == "TOK")
                                    {

                                        f11 = (double.Parse)(reader["vrijednost"].ToString());
                                    }
                                    else
                                    {

                                        g11 = (double.Parse)(reader["vrijednost"].ToString());
                                    }

                                    //e7 = (reader["Proizvedeno_do"].ToString());

                                }
                            }

                            if (kupac.Contains("Technologies"))  // wupertal
                            {

                                if (vrstap.Contains("Valj"))  // wupertal
                                {
                                    if (vo == "TOK")
                                    {

                                        f12 = (double.Parse)(reader["vrijednost"].ToString());
                                    }
                                    else
                                    {

                                        g12 = (double.Parse)(reader["vrijednost"].ToString());
                                    }

                                    //e7 = (reader["Proizvedeno_do"].ToString());
                                    ukupk = ukupk + (double.Parse)(reader["kolicina"].ToString());

                                }
                            }

                            if ((kupac.Contains("KYSUCE")) && (vrstap.Contains("Ku")))  // 
                            {

                                if (vo == "TOK")
                                {

                                    f13 = (double.Parse)(reader["vrijednost"].ToString());
                                }
                                else
                                {

                                    g13 = (double.Parse)(reader["vrijednost"].ToString());
                                }

                                //e7 = (reader["Proizvedeno_do"].ToString());

                            }

                            if ((kupac.Contains("KYSUCE")) && (vrstap.Contains("Prsten")))  // 
                            {
                                if (vo == "TOK")
                                {

                                    f17 = (double.Parse)(reader["vrijednost"].ToString());
                                }
                                else
                                {

                                    g17 = (double.Parse)(reader["vrijednost"].ToString());
                                }

                                //e7 = (reader["Proizvedeno_do"].ToString());

                            }


                            if (kupac.Contains("FERRO"))  // wupertal
                            {
                                if (vo == "TOK")
                                {

                                    f14 = (double.Parse)(reader["vrijednost"].ToString());
                                }
                                else
                                {

                                    g14 = (double.Parse)(reader["vrijednost"].ToString());
                                }

                                //e7 = (reader["Proizvedeno_do"].ToString());

                            }

                            if (kupac.Contains("SONA BLW"))
                            {
                                if (vo == "TOK")
                                {

                                    f15 = (double.Parse)(reader["vrijednost"].ToString());
                                }
                                else
                                {

                                    g15 = (double.Parse)(reader["vrijednost"].ToString());
                                }

                                //e7 = (reader["Proizvedeno_do"].ToString());

                            }

                            if (kupac.Contains("SIGMA"))  // wupertal
                            {
                                if (vo == "TOK")
                                {

                                    f16 = (double.Parse)(reader["vrijednost"].ToString());
                                }
                                else
                                {

                                    g16 = (double.Parse)(reader["vrijednost"].ToString());
                                }

                                //e7 = (reader["Proizvedeno_do"].ToString());


                            }

                            if (kupac.Contains("Brasil"))  // Brasil
                            {
                                if (vo == "TOK")
                                {

                                    f19 = (double.Parse)(reader["vrijednost"].ToString());
                                }
                                else
                                {

                                    g19 = (double.Parse)(reader["vrijednost"].ToString());
                                }

                                //e7 = (reader["Proizvedeno_do"].ToString());

                            }

                            if (kupac.Contains("TECHNOLO"))  // eltmann
                            {
                                if (vo == "TOK")
                                {

                                    f18 = (double.Parse)(reader["vrijednost"].ToString());
                                }
                                else
                                {

                                    g18 = (double.Parse)(reader["vrijednost"].ToString());
                                }

                            }


                            if (kupac.Contains("NEUWEG"))  // NSK


                            {
                                if (vo == "TOK")
                                {

                                    f21 = (double.Parse)(reader["vrijednost"].ToString());
                                }
                                else
                                {

                                    g21 = (double.Parse)(reader["vrijednost"].ToString());
                                }

                                //e7 = (reader["Proizvedeno_do"].ToString());


                            }
                        }
                    }

                }

                string b7 = "", b8 = "", b9 = "", b10 = "", b11 = "", b12 = "", b13 = "", b14 = "", b15 = "", b151= "" , b16 = "", b17 = "", b18 = "", b19 = "", b21 = "", b23="";
                string d70 = "", d80 = "", d90 = "", d100 = "", d110 = "", d120 = "", d130 = "", d140 = "", d150 = "", d1501="" , d160 = "", d170 = "", d180 = "", d190 = "", d210 = "",d230="";
                string hala = "";
                // kupac, planirano
                using (SqlConnection cn = new SqlConnection(connectionString))
                {
                    cn.Open();
                    //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                    // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                    if (ws <= 3)    // d,t,m
                    {
                        sql1 = "rfind.dbo.realizacija '" + dat1 + "','" + dat2 + "',13";
                    }
                    else
                    {
                        sql1 = "rfind.dbo.realizacija '" + dat1 + "','" + dat2 + "',131";
                    }
                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string kupac, idpar,vnar = "";

                    while (reader.Read())
                    {

                        if (ws >= 4)
                        {
                            hala = reader["hala"].ToString();
                            if ((ws == 4) && hala.Contains("3")) // pušta samo halu 1
                            {
                                continue;
                            }
                            if ((ws == 5) && hala.Contains("1"))  // pušta samo 3
                            {
                                continue;
                            }
                        }

                        kupac = reader["Kupac"].ToString();
                        idpar = reader["id_par"].ToString();
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

                            if (vnar.Contains("Valj"))  // wupertal
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

                        if (kupac.Contains("SONA" ))
                        {

                            if (idpar == "121301")
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
                            else
                            {
                                if (vo == "T")
                                {
                                    b151 = (reader["norma"].ToString());
                                }
                                else
                                {
                                    d1501 = (reader["norma"].ToString());
                                }
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

                        if (kupac.Contains("JLG"))  // JLG
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

                        if (kupac.Contains("xxx"))  // 
                        {
                            if (vo == "T")
                            {
                                b23 = (reader["norma"].ToString());
                            }
                            else
                            {
                                d230 = (reader["norma"].ToString());
                            }

                        }

                    }
                    cn.Close();

                }

                #region             // kupac, broj linija koje ne rade
                int n7 = 0, n8 = 0, n9 = 0, n10 = 0, n11 = 0, n12 = 0, n13 = 0, n14 = 0, n15 = 0, n151 = 0,n16 = 0, n17 = 0, n18 = 0, n19 = 0, n20 = 0, n21 = 0, n22=0 ;
                using (SqlConnection cn = new SqlConnection(connectionString))
                {
                    cn.Open();
                    //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                    // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);

                    if (ws<=3)
                        sql1 = "rfind.dbo.ldp_recalc '" + dat1 + "','" + dat2 + "',61";
                    else
                        sql1 = "rfind.dbo.ldp_recalc '" + dat1 + "','" + dat2 + "',610";

                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string kupac,idpar, vnar = "";

                    while (reader.Read())
                    {
                        kupac = reader["Kupac"].ToString();
                        idpar = reader["id_par"].ToString();
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

                            if (vnar.Contains("Valj"))  // wupertal
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
                            if (idpar == "121301")
                            {
                                n15 = (int.Parse)(reader["broj_linija"].ToString());
                            }
                            else
                            {
                                n151 = (int.Parse)(reader["broj_linija"].ToString());
                            }
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
                    i7 = 0.0; i8 = 0.0; i9 = 0.0; i10 = 0.0; i11 = 0.0; i12 = 0.0; i13 = 0.0; i14 = 0.0; i15 = 0.0 ; i151 = 0.0;  i16 = 0.0; i17=0.0 ; i19 = 0.0;i18 = 0.0; i21 = 0.0;i22 = 0;i23 = 0; // 
                    cn.Open();
                    //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                    // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                    if (ws <= 3)   //d,t,m
                    {
                        sql1 = "rfind.dbo.realizacija '" + dat1 + "','" + dat2 + "',2";  // za sve hale
                    }
                    else
                    {
                        sql1 = "rfind.dbo.realizacija '" + dat1 + "','" + dat2 + "',20";  // po halama
                    }
                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string kupac, idpar,vnar = "";

                    while (reader.Read())
                    {

                        if (ws >= 4)
                        {
                            hala = reader["hala"].ToString();
                            if ((ws == 4) && hala.Contains("3")) // pušta samo halu 1
                            {
                                continue;
                            }
                            if ((ws == 5) && hala.Contains("1"))  // pušta samo 3
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
                        idpar = reader["id_par"].ToString();
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

                            if (vnar.Contains("Valj"))  // wupertal
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

                        if (kupac.Contains("JLG"))  // JLG
                        {
                                                       
                                i18 = (double.Parse)(reader["broj_linija"].ToString());                            
                        }



                        if (kupac.Contains("SONA BLW"))
                        {
                            if (idpar == "121301")
                            {
                                i15 = (double.Parse)(reader["broj_linija"].ToString());
                            }
                            else
                            {
                                i151 = (double.Parse)(reader["broj_linija"].ToString());
                            }
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

                        if (kupac.Contains("xxx"))  // 
                        {
                            i23 = (double.Parse)(reader["broj_linija"].ToString());  //
                        }


                        if (ws<=4)
                        {
                            i22=6;
                        }

                    }
                    cn.Close();

                }

                if (1 == 2)   // stara verzija
                {
                    // kupac, korekcija broja linija
                    using (SqlConnection cn = new SqlConnection(connectionString))
                    {
                        i7 = 0.0; i8 = 0.0; i9 = 0.0; i10 = 0.0; i11 = 0.0; i12 = 0.0; i13 = 0.0; i14 = 0.0; i15 = 0.0; i16 = 0.0; i17 = 0.0; i19 = 0.0; i18 = 0.0; i21 = 0.0; i22 = 0; // 
                        cn.Open();
                        //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                        // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                        if (ws <= 0)
                        {
                            //sql1 = "rfind.dbo.realizacija '" + dat1 + "','" + dat2 + "',22";
                        }
                        else
                        {
                            sql1 = "rfind.dbo.realizacija '" + dat1 + "','" + dat2 + "',221";
                        }
                        SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                        SqlDataReader reader = sqlCommand.ExecuteReader();
                        string kupac, vnar = "", Linijanaziv = "";
                        double ukupno = 0.0;

                        while (reader.Read())
                        {

                            if (ws >= 4)
                            {
                                hala = reader["hala"].ToString();
                                if ((ws == 4) && hala.Contains("3")) // pušta samo halu 1
                                {
                                    continue;
                                }
                                if ((ws == 5) && hala.Contains("1"))  // pušta samo 3
                                {
                                    continue;
                                }
                            }
                            kupac = reader["Kupac"].ToString();

                            Linijanaziv = reader["naziv"].ToString();

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
                                if (ws == 3)
                                {
                                    int hj = 0;
                                }
                                i7 = i7  + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                            }

                            if (kupac.Contains("SCHWEINFURT"))
                            {

                                i8 = i8  + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                            }

                            if (kupac.Contains("FAG"))
                            {
                                i9 = i9  + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                            }

                            if (kupac.Contains("ROMANIA"))
                            {
                                i10 = i10  + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                            }

                            if (kupac.Contains("Technologies"))  // wupertal
                            {

                                i11 = i11  + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                            }

                            if (kupac.Contains("KYSUCE"))  // wupertal
                            {
                                i13 = i13  + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                            }

                            if (kupac.Contains("FERRO"))  // wupertal
                            {

                                i14 = i14  + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                            }

                            if (kupac.Contains("SONA BLW"))
                            {
                                i15 = i15  + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                            }

                            if (kupac.Contains("SIGMA"))  // wupertal
                            {
                                i16 = i16  + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                            }

                            if (kupac.Contains("TECHNOLOGIES"))  // eltman
                            {
                                i18 = i18  + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                            }

                            if (kupac.Contains("Brasil"))  // wupertal
                            {
                                i19 = i19  + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                            }

                            if (kupac.Contains("NEUWEG"))  // NSK
                            {
                                i21 = i21  + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                            }

                        }
                        cn.Close();

                    }
                }


                if (ws <= 5)   // nova verzija
                {
                    // kupac, korekcija broja linija
                    using (SqlConnection cn = new SqlConnection(connectionString))
                    {
                        i7 = 0.0; i8 = 0.0; i9 = 0.0; i10 = 0.0; i11 = 0.0; i12 = 0.0; i13 = 0.0; i14 = 0.0; i15 = 0.0; i151 = 0.0;  i16 = 0.0; i17 = 0.0; i19 = 0.0; i18 = 0.0; i21 = 0.0; i23 = 0.0; i22 = 1; // 
                        if (ws <4)  // sve hale
                        {
                            i22 = 69;
                        }
                        if (ws == 4)   // hala 1
                        {
                            i22 = 41;
                        }
                        if (ws == 5)   // hala 3
                        {
                            i22 = 28;
                        }


                        cn.Open();
                        //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                        // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                        if (ws <= 0)
                        {
                            //sql1 = "rfind.dbo.realizacija '" + dat1 + "','" + dat2 + "',22";
                        }
                        else
                        {
                            sql1 = "rfind.dbo.realizacija '" + dat1 + "','" + dat2 + "',221";
                        }
                        SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                        SqlDataReader reader = sqlCommand.ExecuteReader();
                        string kupac,idpar, vnar = "", Linijanaziv = "";
                        double ukupno = 0.0;
                        int red11 = 1;

                        while (reader.Read())
                        {

                            hala = reader["hala"].ToString();
                            if (ws >= 4)
                            {
                                
                                if ((ws == 4) && hala.Contains("3")) // pušta samo halu 1
                                {
                                    continue;
                                }
                                if ((ws == 5) && hala.Contains("1"))  // pušta samo 3
                                {
                                    continue;
                                }
                            }
                            kupac = reader["Kupac"].ToString();
                            idpar = reader["id_par"].ToString();

                            Linijanaziv = reader["parent_pli"].ToString();

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
                                if (ws == 3)
                                {
                                    int hj = 0;
                                }
                                
                                double var1= (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                                i7 = i7 + var1 ;
                                Console.WriteLine("Linija" + Linijanaziv+" udio "+var1.ToString()+"i7 "+i7.ToString());
                                //if (ws == 3)
                                //{
                                //    works10.Cells[red11, 1].Value = Linijanaziv;
                                //    works10.Cells[red11, 2].Value = hala;
                                //    works10.Cells[red11, 3].Value = var1;
                                //    works10.Cells[red11, 4].Value = i7;
                                //    works10.Cells[red11, 5].Value = ws;
                                //    red11++;
                                //}

                            }

                            if (kupac.Contains("SCHWEINFURT"))
                            {

                                i8 = i8 + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                            }

                            if (kupac.Contains("FAG"))
                            {
                                i9 = i9 + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                            }

                            if (kupac.Contains("ROMANIA"))
                            {
                                i10 = i10  + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                            }

                            if (kupac.Contains("Technologies"))  // wupertal
                            {

                                i11 = i11  + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                            }

                            if (kupac.Contains("KYSUCE"))  // wupertal
                            {
                                i13 = i13  + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                            }

                            if (kupac.Contains("FERRO"))  // wupertal
                            {

                                i14 = i14  + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                            }

                            if (kupac.Contains("SONA"))
                            {
                                if (idpar == "121301")
                                {
                                    i15 = i15 + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                                }
                                else
                                {
                                    i151 = i151 + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                                }
                            }

                            if (kupac.Contains("SONA A"))
                            {
                                
                            }

                            if (kupac.Contains("SIGMA"))  // wupertal
                            {
                                i16 = i16 + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                            }

                            if (kupac.Contains("JLG"))  // 
                            {
                                i18 = i18  + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                            }

                            if (kupac.Contains("Brasil"))  // wupertal
                            {
                                i19 = i19  + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                            }

                            if (kupac.Contains("NEUWEG"))  // NSK
                            {
                                i21 = i21  + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                            }

                            if (kupac.Contains("xxx"))  // JLG
                            {
                                i23 = i23 + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                            }


                        }
                        cn.Close();

                    }
                }

                if (ws==3)
                {
                    int zz = 0;
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
                                k7 = (int.Parse)(reader["komada"].ToString());
                            }

                            if (kupac.Contains("SCHWEINFURT"))  // 
                            {
                                k8 = (int.Parse)(reader["komada"].ToString());
                            }

                            if (kupac.Contains("FAG"))
                            {
                                k9 = (int.Parse)(reader["komada"].ToString());
                            }

                            if (kupac.Contains("ROMANIA"))
                            {
                                k10 = (int.Parse)(reader["komada"].ToString());
                            }

                            if (kupac.Contains("Technologies"))  // wupertal
                            {
                                if (vnar.Contains("Prsten"))
                                    k11 = (int.Parse)(reader["komada"].ToString());
                                else
                                    k12 = (int.Parse)(reader["komada"].ToString());
                            }

                            if (kupac.Contains("KYSUCE"))  // 
                            {
                                if (vnar.Contains("Ku"))   // kućište
                                    k13 = (int.Parse)(reader["komada"].ToString());
                                else
                                    k17 = (int.Parse)(reader["komada"].ToString());                            
                            }


                            if (kupac.Contains("FERRO"))  // wupertal
                            {
                                k14 = (int.Parse)(reader["komada"].ToString());
                            }

                            if (kupac.Contains("SONA BLW"))
                            {
                                k15 = (int.Parse)(reader["komada"].ToString());
                            }

                            if (kupac.Contains("SIGMA"))  // wupertal
                            {
                                k16 = (int.Parse)(reader["komada"].ToString());
                            }
                            if (kupac.Contains("ELTMAN"))  // wupertal
                            {
                                k18 = (int.Parse)(reader["komada"].ToString());
                            }
                            if (kupac.Contains("Brasil"))  // wupertal
                            {
                                k19 = (int.Parse)(reader["komada"].ToString());
                            }
                            if (kupac.Contains("NEUWEG"))  // wupertal
                            {
                                k21 = (int.Parse)(reader["komada"].ToString());
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

                        if (kupac.Contains("121302"))   // Sona R
                        {
                            k151 = (double.Parse)(reader["komada"].ToString());
                            q151 = (double.Parse)(reader["vrijednost"].ToString());
                        }


                        if (kupac.Contains("45150"))  // Sigma
                        {
                            k16 = (double.Parse)(reader["komada"].ToString());
                            q16 = (double.Parse)(reader["vrijednost"].ToString());
                        }
                        if (kupac.Contains("221463"))  // JLG
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
                        if (kupac.Contains("xxxx"))  // JLG
                        {
                            k23 = (double.Parse)(reader["komada"].ToString());
                            q23 = (double.Parse)(reader["vrijednost"].ToString());
                        }
                    }
                    cn.Close();
                }

                Console.WriteLine("Izračunato stanje Gotove robe trenutno vrijeme " + DateTime.Now);
                // kupac, škart
                double e7tk = 0.00, e7ts = 0.00, post7t = 0.00, e13tk = 0.00, e13ts = 0.00, post13t = 0.00, e8tk = 0.00, e8ts = 0.00, post8t = 0.00, e9tk = 0.00, e9ts = 0.00, post9t = 0.00, e10tk = 0.00, e10ts = 0.00, post10t = 0.00, e21tk = 0.00, e21ts = 0.0, post21t = 0, post23t =0.0;
                double e11tk = 0.00, e11ts = 0.00, post11t = 0.00, e12tk = 0.00, e12ts = 0.00, post12t = 0.00, e14tk = 0.00, e14ts = 0.00, post14t = 0.00, e15tk = 0.00, e151tk = 0.0, e15ts = 0.00, e151ts = 0.00, post15t = 0.00, post151t = 0.00, e16tk = 0.00, e16ts = 0.00, post16t = 0.00, e18tk = 0.0, e18ts = 0.0, e19tk = 0.0, e23tk = 0.0, post18t=0.0,post18tm=0.0, e19ts = 0.0,  post19t = 0.0 ;
                double post7tm = 0.00, post8tm = 0.00, post9tm = 0.00, post10tm = 0.00, post11tm = 0.00, post12tm = 0.00, post13tm = 0.00, post14tm = 0.00, post15tm = 0.00, post151tm=0.0, post16tm = 0.00, post19tm = 0.00, post21tm = 0.00;
                double e7to = 0, e8to = 0, e9to = 0, e10to = 0, e11to = 0, e12to = 0, e13to = 0, e14to = 0, e15to = 0, e151to =0 , e16to = 0, e17to = 0, e18to = 0, e19to = 0, e21to = 0, e23to=0;
                double e17tk = 0.0, e17ts = 0.0, e23ts=0.0 , post17t = 0.0, post17tm = 0.0,  post23tm=0.0 ; 

                // škart

                using (SqlConnection cn = new SqlConnection(connectionString))
                {
                    cn.Open();
                    //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                    // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                    if (ws <= 3)
                    {
                        sql1 = "rfind.dbo.realizacija '" + dat1 + "','" + dat2 + "',3";
                    }
                    else
                    {
                        sql1 = "rfind.dbo.realizacija '" + dat1 + "','" + dat2 + "',31";
                    }
                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string kupac, idpar,turm, hala1="";

                    while (reader.Read())
                    {

                        if (ws >= 4)
                        {
                            hala1 = reader["hala"].ToString();
                            if ((ws == 4) && hala1.Contains("3")) // pušta samo halu 1
                            {
                                continue;
                            }
                            if ((ws == 5) && hala1.Contains("1"))  // pušta samo 3
                            {
                                continue;
                            }
                        }

                        kupac = reader["Kupac"].ToString();
                        idpar = reader["id_par"].ToString();
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

                            if (kupac.Contains("Valj"))  // wupertal
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

                        if (kupac.Contains("SONA BLW"))
                        {
                            if (idpar == "121301")
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
                            else
                            {
                                if (turm.Contains("Da"))
                                {
                                    e151tk = (double.Parse)(reader["kolicinaok"].ToString());
                                    e151ts = (double.Parse)(reader["otpadobrada"].ToString());
                                    e151to = e151ts;
                                    if (e151tk > 0)
                                        post151t = post151t + e151ts / (e151tk + e151ts);

                                    e151ts = (double.Parse)(reader["otpadmat"].ToString());
                                    if (e151tk > 0)
                                        post151tm = post151tm + e151ts / (e151tk + e151ts);

                                }
                                else
                                {
                                    e151tk = (double.Parse)(reader["kolicinaok"].ToString());
                                    e151ts = (double.Parse)(reader["otpadobrada"].ToString());
                                    e151to = e151ts;
                                    if (e151tk > 0)
                                        post151t = post151t + e151ts / (e151tk + e151ts);

                                    e151ts = (double.Parse)(reader["otpadmat"].ToString());
                                    if (e151tk > 0)
                                        post151tm = post151tm + e151ts / (e151tk + e151ts);

                                }
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


                        if (kupac.Contains("JLG"))  // JLG
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

                        if (kupac.Contains("xxx"))  // 
                        {
                            if (turm.Contains("Da"))
                            {
                                e23tk = (double.Parse)(reader["kolicinaok"].ToString());
                                e23ts = (double.Parse)(reader["otpadobrada"].ToString());
                                e23to = e23ts;
                                if (e23tk > 0)
                                    post23t = post23t + e23ts / (e23tk + e23ts);

                                e23ts = (double.Parse)(reader["otpadmat"].ToString());
                                if (e23tk > 0)
                                    post23tm = post23tm + e23ts / (e23tk + e23ts);

                            }
                            else
                            {
                                e23tk = (double.Parse)(reader["kolicinaok"].ToString());
                                e23ts = (double.Parse)(reader["otpadobrada"].ToString());
                                e23to = e23ts;
                                if (e23tk > 0)
                                    post23t = post23t + e23ts / (e23tk + e23ts);

                                e23ts = (double.Parse)(reader["otpadmat"].ToString());
                                if (e23tk > 0)
                                    post23tm = post23tm + e23ts / (e23ts + e23tk);

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
                // Kaliona po ldp  ??????????????????????????????????
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

                // priprema za dpr
                string kkal7 = "0", kkal8 = "0", kkal9 = "0", kkal10 = "0", kkal11 = "0", kkal12 = "0", kkal13 = "0", kkal14 = "0", kkal15 = "0", kkal151 ="0", kkal16 = "0", kkal17 = "0", kkal18 = "0", kkal19 = "0", kkal21 = "0" , kkal23 = "0";
                string vkal7 = "0", vkal8 = "0", vkal9 = "0", vkal10 = "0", vkal11 = "0", vkal12 = "0", vkal13 = "0", vkal14 = "0", vkal15 = "0", vkal151 ="0", vkal16 = "0", vkal17 = "0", vkal18 = "0", vkal19 = "0", vkal21 = "0" , vkal23 = "0";
                using (SqlConnection cn = new SqlConnection(connectionString))
                {
                    cn.Open();
                    //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                    // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                    sql1 = "rfind.dbo.realizacija '" + dat1 + "','" + dat2 + "',61";
                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string pec;

                    while (reader.Read())
                    {
                        //pec = reader["pec"].ToString();
                        //if (reader["pec"] != DBNull.Value)
                        //{
                        //    if (pec.Contains("CODERE"))
                        //    {
                        //        b312 = (double.Parse)(reader["tezina"].ToString());
                        //    }
                        //    else
                        //    {
                        //        if (reader["tezina"] != DBNull.Value)
                        //        {
                        //            e312 = (double.Parse)(reader["tezina"].ToString());
                        //        }
                        //    }
                        //}
                        string partner1 = reader["id_par"].ToString();
                        if (partner1 == "274")
                        {
                            kkal7 = reader["kolicinaok"].ToString();
                            vkal7 = reader["tezinaprokom"].ToString().Replace(',', '.');
                        }
                        if (partner1 == "121269")
                        {
                            kkal8 = reader["kolicinaok"].ToString();
                            vkal8 = reader["tezinaprokom"].ToString().Replace(',', '.');
                        }
                        if (partner1 == "221452")
                        {
                            kkal9 = reader["kolicinaok"].ToString();
                            vkal9 = reader["tezinaprokom"].ToString().Replace(',', '.');
                        }
                        if (partner1 == "221453")
                        {
                            kkal10 = reader["kolicinaok"].ToString();
                            vkal10 = reader["tezinaprokom"].ToString().Replace(',', '.');
                        }
                        if (partner1 == "121274")
                        {
                            kkal11 = reader["kolicinaok"].ToString();
                            vkal11 = reader["tezinaprokom"].ToString().Replace(',', '.');
                        }
                        if ((partner1 == "121274") && (1 == 2))
                        {
                            kkal12 = reader["kolicinaok"].ToString();
                            vkal12 = reader["tezinaprokom"].ToString().Replace(',', '.');
                        }
                        if (partner1 == "121273")
                        {
                            kkal13 = reader["kolicinaok"].ToString();
                            vkal13 = reader["tezinaprokom"].ToString().Replace(',', '.');
                        }

                        if (partner1 == "8752")
                        {
                            kkal14 = reader["kolicinaok"].ToString();
                            vkal14 = reader["tezinaprokom"].ToString().Replace(',', '.');
                        }
                        if (partner1 == "121301")
                        {
                            kkal15 = reader["kolicinaok"].ToString();
                            vkal15 = reader["tezinaprokom"].ToString().Replace(',', '.');
                        }
                        if (partner1 == "121302")
                        {
                            kkal151 = reader["kolicinaok"].ToString();
                            vkal151 = reader["tezinaprokom"].ToString().Replace(',', '.');
                        }

                        if (partner1 == "45150")
                        {
                            kkal16 = reader["kolicinaok"].ToString();
                            vkal16 = reader["tezinaprokom"].ToString().Replace(',', '.');
                        }
                        if (partner1 == "121273")
                        {
                            kkal17 = reader["kolicinaok"].ToString();
                            vkal17 = reader["tezinaprokom"].ToString().Replace(',', '.');
                        }
                        if (partner1 == "221463")  // JLG
                        {
                            kkal18 = reader["kolicinaok"].ToString();
                            vkal18 = reader["tezinaprokom"].ToString().Replace(',', '.');
                        }
                        if (partner1 == "121278")
                        {
                            kkal19 = reader["kolicinaok"].ToString();
                            vkal19 = reader["tezinaprokom"].ToString().Replace(',', '.');
                        }
                        if (partner1 == "121400")
                        {
                            kkal21 = reader["kolicinaok"].ToString();
                            vkal21 = reader["tezinaprokom"].ToString().Replace(',', '.');
                        }
                        if (partner1 == "xxxx")   // 
                        {
                            kkal21 = reader["kolicinaok"].ToString();
                            vkal21 = reader["tezinaprokom"].ToString().Replace(',', '.');
                        }



                    }
                    cn.Close();

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

                Worksheet worksheet = workbook.Worksheets.Item[ws] as Worksheet;
                if (worksheet == null)
                    return;

                int ul24 = 0, ul15 = 0, iz24 = 0, iz15 = 0;
                int naBolovanju = 0, naGO = 0, ostalo = 0;
                if (ws != 2 && ws!=3 ) 
                {

                    // pregled izostanka, zbirno

                    using (SqlConnection cn3 = new SqlConnection(connectionString2))
                    {
                        cn3.Open();
                        //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                        // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);

                        sql1 = "rfind.dbo.ldp_recalc2 '" + dat1 + "','" + dat1 + "',7";
                        SqlCommand sqlCommand = new SqlCommand(sql1, cn3);

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

                        cn3.Close();
                    }

                    // koji nisu došli mo mjestu troška
                    using (SqlConnection cn = new SqlConnection(connectionString2))
                    {
                        cn.Open();
                        //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                        // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                        if (ws<=3)
                        {
                            sql1 = "rfind.dbo.ldp_recalc2 '" + datLDP + "','" + datLDP + "',72";
                        }
                        else
                        {
                            sql1 = "rfind.dbo.ldp_recalc2 '" + datLDP + "','" + datLDP + "',720";
                        }
                        
                        SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                        SqlDataReader reader = sqlCommand.ExecuteReader();
                        //worksizostanci.Rows.Cells[4, 3].value = worksizostanci.Rows.Cells[4, 3].value + " " + datreps;
                        string vrstao = "", mt = "",hala1="";
                        int broj = 0;

                        int i = 0;
                        while (reader.Read())
                        {
                            if (ws >= 4)
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

                                if ((ws == 4) && hala1.Contains("3")) // pušta samo halu 1
                                {
                                    continue;
                                }
                                if ((ws == 5) && hala1.Contains("1"))  // pušta samo 3
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
                                    worksheet.Rows.Cells[10, 21].value = broj;
                                }
                                if (vrstao.Contains("Godi"))
                                {
                                    worksheet.Rows.Cells[10, 22].value = broj;
                                }
                                if (vrstao.Contains("Nije"))
                                {
                                    worksheet.Rows.Cells[10, 23].value = broj;
                                }

                            }
                            if (mt.Contains("Kaljenje"))
                            {
                                if (vrstao.Contains("Bolovanje"))
                                {
                                    worksheet.Rows.Cells[11, 21].value = broj;
                                }
                                if (vrstao.Contains("Godi"))
                                {
                                    worksheet.Rows.Cells[11, 22].value = broj;
                                }
                                if (vrstao.Contains("Nije"))
                                {
                                    worksheet.Rows.Cells[11, 23].value = broj;
                                }

                            }
                            if (mt.Contains("Alatnica"))
                            {
                                if (vrstao.Contains("Bolovanje"))
                                {
                                    worksheet.Rows.Cells[12, 21].value = broj;
                                }
                                if (vrstao.Contains("Godi"))
                                {
                                    worksheet.Rows.Cells[12, 22].value = broj;
                                }
                                if (vrstao.Contains("Nije"))
                                {
                                    worksheet.Rows.Cells[12, 23].value = broj;
                                }

                            }
                            if (mt.Contains("Održavanje"))
                            {
                                if (vrstao.Contains("Bolovanje"))
                                {
                                    worksheet.Rows.Cells[13, 21].value = broj;
                                }
                                if (vrstao.Contains("Godi"))
                                {
                                    worksheet.Rows.Cells[13, 22].value = broj;
                                }
                                if (vrstao.Contains("Nije"))
                                {
                                    worksheet.Rows.Cells[13, 23].value = broj;
                                }

                            }
                            if (mt.Contains("Šteleri"))
                            {
                                if (vrstao.Contains("Bolovanje"))
                                {
                                    worksheet.Rows.Cells[14, 21].value = broj;
                                }
                                if (vrstao.Contains("Godi"))
                                {
                                    worksheet.Rows.Cells[14, 22].value = broj;
                                }
                                if (vrstao.Contains("Nije"))
                                {
                                    worksheet.Rows.Cells[14, 23].value = broj;
                                }

                            }
                            if (mt.Contains("Kvaliteta"))
                            {
                                if (vrstao.Contains("Bolovanje"))
                                {
                                    worksheet.Rows.Cells[15, 21].value = broj;
                                }
                                if (vrstao.Contains("Godi"))
                                {
                                    worksheet.Rows.Cells[15, 22].value = broj;
                                }
                                if (vrstao.Contains("Nije"))
                                {
                                    worksheet.Rows.Cells[15, 23].value = broj;
                                }

                            }
                            if (mt.Contains("Zaštita"))
                            {
                                if (vrstao.Contains("Bolovanje"))
                                {
                                    worksheet.Rows.Cells[16, 21].value = broj;
                                }
                                if (vrstao.Contains("Godi"))
                                {
                                    worksheet.Rows.Cells[16, 22].value = broj;
                                }
                                if (vrstao.Contains("Nije"))
                                {
                                    worksheet.Rows.Cells[16, 23].value = broj;
                                }
                            }
                        }
                        i++;

                        cn.Close();
                    }
                    Console.WriteLine("Izostanci po mjestu troška trenutno vrijeme " + DateTime.Now);

                    // broj pristunih po MT

                    using (SqlConnection cn = new SqlConnection(connectionString2))
                    {
                        cn.Open();
                        //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                        // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                        if (ws <= 3)
                        {
                            sql1 = "rfind.dbo.ldp_recalc2 '" + datLDP + "','" + datLDP + "',73";
                        }
                        else
                        {
                            sql1 = "rfind.dbo.ldp_recalc2 '" + datLDP + "','" + datLDP + "',730";
                        }
                        SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                        SqlDataReader reader = sqlCommand.ExecuteReader();
                        //worksizostanci.Rows.Cells[4, 3].value = worksizostanci.Rows.Cells[4, 3].value + " " + datreps;
                        string vrstao = "", mt = "",hala1="";
                        int broj = 0;

                        int i = 0;
                        while (reader.Read())
                        {
                            if (ws >= 4)
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

                                if ((ws == 4) && hala1.Contains("3")) // pušta samo halu 1
                                {
                                    continue;
                                }
                                if ((ws == 5) && hala1.Contains("1"))  // pušta samo 3
                                {
                                    continue;
                                }
                            }

                            mt = reader["mtroska"].ToString();
                            broj = (int.Parse)(reader["broj"].ToString());


                            if (mt.Contains("Tokarenje"))
                            {
                                worksheet.Rows.Cells[10, 24].value = broj;

                            }
                            if (mt.Contains("Kaljenje"))
                            {
                                worksheet.Rows.Cells[11, 24].value = broj;
                            }
                            if (mt.Contains("Alatnica"))
                            {
                                worksheet.Rows.Cells[12, 24].value = broj;
                            }
                            if (mt.Contains("Održavanje"))
                            {
                                worksheet.Rows.Cells[13, 24].value = broj;
                            }
                            if (mt.Contains("Šteleri"))
                            {
                                worksheet.Rows.Cells[14, 24].value = broj;
                            }

                            if (mt.Contains("Kvaliteta"))
                            {
                                worksheet.Rows.Cells[15, 24].value = broj;
                            }
                            if (mt.Contains("Zaštita"))
                            {
                                worksheet.Rows.Cells[16, 24].value = broj;
                            }
                            if (mt.Contains("Ukupno"))
                            {
                                worksheet.Rows.Cells[7, 19].value = broj;
                            }

                        }
                        i++;

                        cn.Close();
                    }
                    Console.WriteLine("Došli po mjestu troška trenutno vrijeme " + DateTime.Now);



                } // ws = 1
                  // planiranje share
                Workbook workbook2 = excel.Workbooks.Open(@"\\192.168.0.3\voditelj pogona\PRIPREMA ZA SASTANAK-PLANIRANJE\Planiranje (SHARE).xlsx", ReadOnly: true, Editable: false);
                // transporti 2017
                //Workbook workbook3 = excel.Workbooks.Open(@"\\192.168.0.3\Referentice$\Transporti 2017.xlsm", ReadOnly: true, Editable: false);
                // Workbook workbook3 = excel.Workbooks.Open(@"c:\brisi\Transporti 2017.xlsx", ReadOnly: true, Editable: false);
                //Workbook workbook2 = excel.Workbooks.Open(@"\\192.168.0.3\voditelj pogona\dokumenti\Planiranje (SHARE)_.xlsx", ReadOnly: true, Editable: false);
                //Worksheet worksheet31 = workbook3.Worksheets.Item[1] as Worksheet;  // 
                // oMissing = null;
                //worksheet31.UsedRange.AutoFilter(1, worksheet31, Microsoft.Office.Interop.Excel.XlAutoFilterOperator.xlAnd, oMissing, false);

                //System.Data.OleDb.OleDbConnection dbConn = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:\\Sample.xlsx;Extended Properties='Excel 12.0 Xml;HDR=YES'");
                //string connectionStringE = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\192.168.0.3\referentice$\Transporti 2017.xlsx;Extended Properties='Excel 12.0 Xml;HDR=YES'";
                string connectionStringE = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=t:\Transporti 2018.xlsx;Extended Properties='Excel 12.0 Xml;HDR=YES'";
                if (ws <=1 && 1==12)   // samo za dnevni i tjedni izvještaj
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
                ////int iRow = dt.Rows.Count;
                ////foreach (DataRow dr in dt.Rows)
                ////    Console.WriteLine(dr[0]);


                ////Range findRng = Fruits.Find("28.7.2017", Missing.Value, XlFindLookIn.xlValues, Missing.Value, Missing.Value, XlSearchDirection.xlNext, false, false, Missing.Value);
                ////vv1 = findRng.Cells[2, 1].value;
                ////vv2 = findRng.Cells[2, 2].value;
                ////vv3 = findRng.Cells[2, 13].value;

                ////findRng = Fruits.FindNext(findRng);
                ////vv1 = findRng.Cells[2, 1].value;
                ////vv2 = findRng.Cells[2, 2].value;
                ////vv3 = findRng.Cells[2, 13].value;

                //Microsoft.Office.Interop.Excel.Range xlFound = Fruits.EntireRow.Find("28.7.2017", Missing.Value, Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues, Microsoft.Office.Interop.Excel.XlLookAt.xlPart, Microsoft.Office.Interop.Excel.XlSearchOrder.xlByColumns, Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext,true, Missing.Value, Missing.Value);
                ////findRng.Select();
                ////int z1 = xlFound.Rows.Count;
                ////foreach(Microsoft.Office.Interop.Excel.Range row in findRng.Rows)
                ////{
                ////    vv1 = row.Cells[2, 1].value;
                ////    vv2 = row.Cells[2, 2].value;
                ////    vv3 = row.Cells[2, 13].value;

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
                //    vv1 = area.Rows.Cells[2, 1].value;
                //}
                ////Microsoft.Office.Interop.Excel.Range areaRange = filteredRange.Areas.get_Item(1);
                ////DateTime vv1 = areaRange.Rows.Cells[2, 1].value;

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

                //double v1 = worksheet1.Rows.Cells[red1, 8].value;
                //worksheet.Rows.Cells[7, 13].value = worksheet1.Rows.Cells[red1, 8].value;  //  BDF naš dnevni kapacitet
                //worksheet.Rows.Cells[7, 12].value = worksheet1.Rows.Cells[red1, 9].value;  //  BDF naš dnevni k. trazi
                //worksheet.Rows.Cells[7, 11].value = worksheet1.Rows.Cells[red1, 16].value;  //  BDF  Lager gotove robe  komada

                //double v1 = worksheet1.Rows.Cells[red1, 8].value;

                worksheet.Rows.Cells[7, 2].value = b7;
                worksheet.Rows.Cells[7, 4].value = d70;
                worksheet.Rows.Cells[8, 2].value = b8;
                worksheet.Rows.Cells[8, 4].value = d80;
                worksheet.Rows.Cells[9, 2].value = b9;
                worksheet.Rows.Cells[9, 4].value = d90;
                worksheet.Rows.Cells[10, 2].value = b10;
                worksheet.Rows.Cells[10, 4].value = d100;
                worksheet.Rows.Cells[11, 2].value = b11;
                worksheet.Rows.Cells[11, 4].value = d110;
                worksheet.Rows.Cells[12, 2].value = b12;
                worksheet.Rows.Cells[12, 4].value = d120;
                worksheet.Rows.Cells[13, 2].value = b13;
                worksheet.Rows.Cells[13, 4].value = d130;
                worksheet.Rows.Cells[14, 2].value = b14;
                worksheet.Rows.Cells[14, 4].value = d140;

                worksheet.Rows.Cells[15, 2].value = b15;
                worksheet.Rows.Cells[15, 4].value = d150;
                worksheet.Rows.Cells[16, 2].value = b151;
                worksheet.Rows.Cells[16, 4].value = d1501;


                worksheet.Rows.Cells[17, 2].value = b16;
                worksheet.Rows.Cells[17, 4].value = d160;

                worksheet.Rows.Cells[18, 2].value = b17;
                worksheet.Rows.Cells[18, 4].value = d170;

                worksheet.Rows.Cells[19, 2].value = b18;
                worksheet.Rows.Cells[19, 4].value = d180;

                worksheet.Rows.Cells[20, 2].value = b19;
                worksheet.Rows.Cells[20, 4].value = d190;

                red1 = 141 + dana;

                
                // Pakiranje, to se više ne radi

                connectionStringE = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=l:\PRIPREMA ZA SASTANAK-PLANIRANJE\Planiranje (SHARE).xlsx;Extended Properties='Excel 12.0 Xml;HDR=YES'";
                //connectionStringE = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=l:\dokumenti\Planiranje (SHARE).xlsx;Extended Properties='Excel 12.0 Xml;HDR=YES'";
                v1 = ""; v2 = ""; v3 = "";
                //using (System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection(connectionStringE))
                //{
                //    System.Data.OleDb.OleDbCommand cmd = new System.Data.OleDb.OleDbCommand();
                //    cmd.Connection = conn;
                //    cmd.CommandText = "SELECT * from [pakiranje$] where day(datum)=" + dan1 + " and month(datum)=" + m1 + " and year(datum)=" + g1;

                //      conn.Open();
                //    System.Data.IDataReader dr = cmd.ExecuteReader();

                //    while (dr.Read())
                //    {

                //        v1 = dr[4].ToString();
                //    }
                //}

                worksheet.Rows.Cells[23, 2].value = v1;  // FAG
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
                //worksheet.Rows.Cells[15, 12].value = v1;  //  sona kupac trazi
                //worksheet.Rows.Cells[15, 13].value = v2;  //  sona  naš kapacitet
                //worksheet.Rows.Cells[15, 2].value = v3;  //  sona planirano
                //worksheet.Rows.Cells[15, 11].value = v4;  //  sona LGR kom.


                red1 = 141 + dana;
                v1 = ""; v2 = ""; v3 = ""; v4 = "";
                if (ws < 4)  // samo dnevni i tjedni izvještaj
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
                    worksheet.Rows.Cells[34, 9].value = v1;  //  Kaliona   cementacija kg
                
                //workbook2.Close(false);
                workbook2.Close(false, Type.Missing, Type.Missing);
                //workbook2.Close(false, Type.Missing, Type.Missing);

                GC.Collect();

                //objExcel.Application.Quit();
                //objExcel.Quit();
                Console.WriteLine("Zatvoreno Planiranje share   trenutno vrijeme " + DateTime.Now);

                // }  // if ws==1
                //Range row1 = worksheet.Rows.Cells[7, 3];
                //Range row2 = worksheet.Rows.Cells[13, 3];

                //row1.Value = c7;
                //row2.Value = c13;

                
                worksheet.Rows.Cells[1, 1].value = worksheet.Rows.Cells[1, 1].value + " " + dat10;

                worksheet.Rows.Cells[16, 2].value = b16;  // planiranje  Sigma
                //worksheet.Rows.Cells[16, 11].value = k16;  // LGR  Sigma

                worksheet.Rows.Cells[20, 2].value = b19;  // planiranje Brasil
                worksheet.Rows.Cells[22, 2].value = b21;  // planiranje nsk

                worksheet.Rows.Cells[7, 3].value = c7;  // količina tok
                worksheet.Rows.Cells[8, 3].value = c8;
                worksheet.Rows.Cells[9, 3].value = c9;
                worksheet.Rows.Cells[10, 3].value = c10;
                worksheet.Rows.Cells[11, 3].value = c11;
                worksheet.Rows.Cells[12, 3].value = c12;
                worksheet.Rows.Cells[13, 3].value = c13;
                worksheet.Rows.Cells[14, 3].value = c14;
                worksheet.Rows.Cells[15, 3].value = c15;
                worksheet.Rows.Cells[16, 3].value = c151;
                worksheet.Rows.Cells[17, 3].value = c16;
                worksheet.Rows.Cells[18, 3].value = c17;
                worksheet.Rows.Cells[19, 3].value = c18;
                worksheet.Rows.Cells[20, 3].value = c19;
                worksheet.Rows.Cells[22, 3].value = c21;
                worksheet.Rows.Cells[21, 3].value = c23;
                worksheet.Rows.Cells[23, 3].value = c23;

                worksheet.Rows.Cells[7, 5].value = e7;  // količina, dodatne operacije
                worksheet.Rows.Cells[8, 5].value = e8;
                worksheet.Rows.Cells[9, 5].value = e9;
                worksheet.Rows.Cells[10, 5].value = e10;
                worksheet.Rows.Cells[11, 5].value = e11;
                worksheet.Rows.Cells[12, 5].value = e12;
                worksheet.Rows.Cells[13, 5].value = e13;
                worksheet.Rows.Cells[14, 5].value = e14;
                worksheet.Rows.Cells[15, 5].value = e15;
                worksheet.Rows.Cells[16, 5].value = e151;
                worksheet.Rows.Cells[17, 5].value = e16;
                worksheet.Rows.Cells[18, 5].value = e17;
                worksheet.Rows.Cells[19, 5].value = e18;
                worksheet.Rows.Cells[20, 5].value = e19;
                worksheet.Rows.Cells[22, 5].value = e21;
                worksheet.Rows.Cells[21, 5].value = e23;
                worksheet.Rows.Cells[23, 5].value = e23;

                worksheet.Rows.Cells[7, 6].value = f7;  // vrijednost tokarenja
                worksheet.Rows.Cells[8, 6].value = f8;
                worksheet.Rows.Cells[9, 6].value = f9;
                worksheet.Rows.Cells[10, 6].value = f10;
                worksheet.Rows.Cells[11, 6].value = f11;
                worksheet.Rows.Cells[12, 6].value = f12;
                worksheet.Rows.Cells[13, 6].value = f13;
                worksheet.Rows.Cells[14, 6].value = f14;
                worksheet.Rows.Cells[15, 6].value = f15;
                worksheet.Rows.Cells[16, 6].value = f151;
                worksheet.Rows.Cells[17, 6].value = f16;
                worksheet.Rows.Cells[18, 6].value = f17;
                worksheet.Rows.Cells[19, 6].value = f18;
                worksheet.Rows.Cells[20, 6].value = f19;                
                worksheet.Rows.Cells[22, 6].value = f21;
                worksheet.Rows.Cells[23, 6].value = f23;


                worksheet.Rows.Cells[7, 7].value = g7;  // vrijednost dodatne operacije
                worksheet.Rows.Cells[8, 7].value = g8;
                worksheet.Rows.Cells[9, 7].value = g9;
                worksheet.Rows.Cells[10, 7].value = g10;
                worksheet.Rows.Cells[11, 7].value = g11;
                worksheet.Rows.Cells[12, 7].value = g12;
                worksheet.Rows.Cells[13, 7].value = g13;
                worksheet.Rows.Cells[14, 7].value = g14;
                worksheet.Rows.Cells[15, 7].value = g15;
                worksheet.Rows.Cells[16, 7].value = g151;
                worksheet.Rows.Cells[17, 7].value = g16;
                worksheet.Rows.Cells[18, 7].value = g17;
                worksheet.Rows.Cells[19, 7].value = g18;
                worksheet.Rows.Cells[20, 7].value = g19;
                worksheet.Rows.Cells[22, 7].value = g21;
                worksheet.Rows.Cells[23, 7].value = g23;

                if (ws == 1)
                {
                 
                    using (SqlConnection cn2 = new SqlConnection(connectionString2))
                    {

                        cn2.Open();
                        sql1                   = "delete from dpr where datum='"+ dat1+"'";
                        SqlCommand sqlCommand1 = new SqlCommand(sql1, cn2);
                        SqlDataReader reader1  = sqlCommand1.ExecuteReader();
                        cn2.Close();

                        cn2.Open();
                        sql1 = "delete from planiranje.dbo.planiranjeshare where datum='" + dat1 + "'";
                        sqlCommand1 = new SqlCommand(sql1, cn2);
                        reader1 = sqlCommand1.ExecuteReader();
                        cn2.Close();


                        cn2.Open();
                        if (b7.TrimEnd() == "") b7 = "0"; if (b8.TrimEnd() == "") b8 = "0"; if (b9.TrimEnd() == "") b9 = "0"; if (b10.TrimEnd() == "") b10 = "0"; if (b11.TrimEnd() == "") b11 = "0"; if (b12.TrimEnd() == "") b12 = "0";
                        if (b13.TrimEnd() == "") b13 = "0"; if (b14.TrimEnd() == "") b14 = "0"; if (b15.TrimEnd() == "") b15 = "0"; if (b151.TrimEnd() == "") b151 = "0"; if (b16.TrimEnd() == "") b16 = "0"; if (b17.TrimEnd() == "") b17 = "0"; if (b18.TrimEnd() == "") b18 = "0"; if (b19.TrimEnd() == "") b19 = "0"; if (b21.TrimEnd() == "") b21 = "0";
                        if (b23.TrimEnd() == "") b23 = "0";

                        if (d70.TrimEnd() == "") d70 = "0"; if (d80.TrimEnd() == "") d80 = "0"; if (d90.TrimEnd() == "") d90 = "0"; if (d100.TrimEnd() == "") d100 = "0"; if (d110.TrimEnd() == "") d110 = "0"; if (d120.TrimEnd() == "") d120 = "0";
                        if (d130.TrimEnd() == "") d130 = "0"; if (d140.TrimEnd() == "") d140 = "0"; if (d150.TrimEnd() == "") d150 = "0"; if (d1501.TrimEnd() == "") d1501 = "0"; if (d160.TrimEnd() == "") d160 = "0"; if (d170.TrimEnd() == "") d170 = "0"; if (d180.TrimEnd() == "") d180 = "0"; if (d190.TrimEnd() == "") d190 = "0"; if (d210.TrimEnd() == "") d210 = "0";
                        if (d230.TrimEnd() == "") d230 = "0";


                        CultureInfo myCI = new CultureInfo("de-DE");
                        Calendar myCal = myCI.Calendar;
                        CalendarWeekRule myCWR = myCI.DateTimeFormat.CalendarWeekRule;
                        DayOfWeek myFirstDOW = myCI.DateTimeFormat.FirstDayOfWeek;

                        string wn1 = myCal.GetWeekOfYear(d3, myCWR, myFirstDOW).ToString();
                         
                        string sada = d3.Year.ToString() + "-" + d3.Month.ToString() + "-"+d3.Day.ToString() + " " + d3.Hour.ToString() + ":" + d3.Minute.ToString() + ":" + d3.Second.ToString();
                        sql1 = "insert into dpr (kolicina_kal,vrijednost_kal,datum,tjedan,kupac,sgr_kom,sgr_eur,id_kupac,planiranokol_tok,planiranokol_do,kolicina_tok,kolicina_do,vrijednost_tok,vrijednost_do,datumunosa,mmgggg) values ("+ kkal7 + "," + vkal7 + ",'" + dat1 + "'," + wn1 + ",'" + nkupac7 + "'," + ((int)(k7)).ToString() + "," + ((int)(q7)).ToString() + "," + kupac7 + "," + b7 + "," + d70 + "," + c7.ToString() + "," + e7.ToString() + "," + f7.ToString().Replace(',', '.') + "," + g7.ToString().Replace(',', '.') + ",'" + sada + "','" + mmgggg1 + "');" +
                                "insert into dpr (kolicina_kal,vrijednost_kal,datum,tjedan,kupac,sgr_kom,sgr_eur,id_kupac,planiranokol_tok,planiranokol_do,kolicina_tok,kolicina_do,vrijednost_tok,vrijednost_do,datumunosa,mmgggg) values (" + kkal8 + "," + vkal8 + ",'" + dat1 + "'," + wn1 + ",'" + nkupac8 + "'," + ((int)(k8)).ToString() + "," + ((int)(q8)).ToString() + "," + kupac8 + "," + b8 + "," + d80 + "," + c8.ToString() + "," + e8.ToString() + "," + f8.ToString().Replace(',', '.') + "," + g8.ToString().Replace(',', '.') + ",'" + sada + "','" + mmgggg1 + "');" +
                                "insert into dpr (kolicina_kal,vrijednost_kal,datum,tjedan,kupac,sgr_kom,sgr_eur,id_kupac,planiranokol_tok,planiranokol_do,kolicina_tok,kolicina_do,vrijednost_tok,vrijednost_do,datumunosa,mmgggg) values (" + kkal9 + "," + vkal9 + ", '" + dat1 + "'," + wn1 + ",'" + nkupac9 + "'," + ((int)(k9)).ToString() + "," + ((int)(q9)).ToString() + "," + kupac9 + "," + b9 + "," + d90 + "," + c9.ToString() + "," + e9.ToString() + "," + f9.ToString().Replace(',', '.') + "," + g9.ToString().Replace(',', '.') + ",'" + sada + "','" + mmgggg1 + "');" +
                                "insert into dpr (kolicina_kal,vrijednost_kal,datum,tjedan,kupac,sgr_kom,sgr_eur,id_kupac,planiranokol_tok,planiranokol_do,kolicina_tok,kolicina_do,vrijednost_tok,vrijednost_do,datumunosa,mmgggg) values (" + kkal10 + "," + vkal10 + ", '" + dat1 + "'," + wn1 + ",'" + nkupac10 + "'," + ((int)(k10)).ToString() + "," + ((int)(q10)).ToString() + "," + kupac10 + "," + b10 + "," + d100 + "," + c10.ToString() + "," + e10.ToString() + "," + f10.ToString().Replace(',', '.') + "," + g10.ToString().Replace(',', '.') + ",'" + sada + "','" + mmgggg1 + "');" +
                                "insert into dpr (kolicina_kal,vrijednost_kal,datum,tjedan,kupac,sgr_kom,sgr_eur,id_kupac,planiranokol_tok,planiranokol_do,kolicina_tok,kolicina_do,vrijednost_tok,vrijednost_do,datumunosa,mmgggg) values (" + kkal11 + "," + vkal11 + ", '" + dat1 + "'," + wn1 + ",'" + nkupac11 + "'," + ((int)(k11)).ToString() + "," + ((int)(q11)).ToString() + "," + kupac11 + "," + b11 + "," + d110 + "," + c11.ToString() + "," + e11.ToString() + "," + f11.ToString().Replace(',', '.') + "," + g11.ToString().Replace(',', '.') + ",'" + sada + "','" + mmgggg1 + "');" +
                                "insert into dpr (kolicina_kal,vrijednost_kal,datum,tjedan,kupac,sgr_kom,sgr_eur,id_kupac,planiranokol_tok,planiranokol_do,kolicina_tok,kolicina_do,vrijednost_tok,vrijednost_do,datumunosa,mmgggg) values (" + kkal12 + "," + vkal12 + ", '" + dat1 + "'," + wn1 + ",'" + nkupac12 + "'," + ((int)(k12)).ToString() + "," + ((int)(q12)).ToString() + "," + kupac12 + "," + b12 + "," + d120 + "," + c12.ToString() + "," + e12.ToString() + "," + f12.ToString().Replace(',', '.') + "," + g12.ToString().Replace(',', '.') + ",'" + sada + "','" + mmgggg1 + "');" +
                                "insert into dpr (kolicina_kal,vrijednost_kal,datum,tjedan,kupac,sgr_kom,sgr_eur,id_kupac,planiranokol_tok,planiranokol_do,kolicina_tok,kolicina_do,vrijednost_tok,vrijednost_do,datumunosa,mmgggg) values (" + kkal13 + "," + vkal13 + ", '" + dat1 + "'," + wn1 + ",'" + nkupac13 + "'," + ((int)(k13)).ToString() + "," + ((int)(q13)).ToString() + "," + kupac13 + "," + b13 + "," + d130 + "," + c13.ToString() + "," + e13.ToString() + "," + f13.ToString().Replace(',', '.') + "," + g13.ToString().Replace(',', '.') + ",'" + sada + "','" + mmgggg1 + "');" +
                                "insert into dpr (kolicina_kal,vrijednost_kal,datum,tjedan,kupac,sgr_kom,sgr_eur,id_kupac,planiranokol_tok,planiranokol_do,kolicina_tok,kolicina_do,vrijednost_tok,vrijednost_do,datumunosa,mmgggg) values (" + kkal14 + "," + vkal14 + ", '" + dat1 + "'," + wn1 + ",'" + nkupac14 + "'," + ((int)(k14)).ToString() + "," + ((int)(q14)).ToString() + "," + kupac14 + "," + b14 + "," + d140 + "," + c14.ToString() + "," + e14.ToString() + "," + f14.ToString().Replace(',', '.') + "," + g14.ToString().Replace(',', '.') + ",'" + sada + "','" + mmgggg1 + "');" +
                                "insert into dpr (kolicina_kal,vrijednost_kal,datum,tjedan,kupac,sgr_kom,sgr_eur,id_kupac,planiranokol_tok,planiranokol_do,kolicina_tok,kolicina_do,vrijednost_tok,vrijednost_do,datumunosa,mmgggg) values (" + kkal15 + "," + vkal15 + ", '" + dat1 + "'," + wn1 + ",'" + nkupac15 + "'," + ((int)(k15)).ToString() + "," + ((int)(q15)).ToString() + "," + kupac15 + "," + b15 + "," + d150 + "," + c15.ToString() + "," + e15.ToString() + "," + f15.ToString().Replace(',', '.') + "," + g15.ToString().Replace(',', '.') + ",'" + sada + "','" + mmgggg1 + "');" +
                                "insert into dpr (kolicina_kal,vrijednost_kal,datum,tjedan,kupac,sgr_kom,sgr_eur,id_kupac,planiranokol_tok,planiranokol_do,kolicina_tok,kolicina_do,vrijednost_tok,vrijednost_do,datumunosa,mmgggg) values (" + kkal151 + "," + vkal151 + ", '" + dat1 + "'," + wn1 + ",'" + nkupac151 + "'," + ((int)(k151)).ToString() + "," + ((int)(q151)).ToString() + "," + kupac151 + "," + b151 + "," + d1501 + "," + c151.ToString() + "," + e151.ToString() + "," + f151.ToString().Replace(',', '.') + "," + g151.ToString().Replace(',', '.') + ",'" + sada + "','" + mmgggg1 + "');" +
                                "insert into dpr (kolicina_kal,vrijednost_kal,datum,tjedan,kupac,sgr_kom,sgr_eur,id_kupac,planiranokol_tok,planiranokol_do,kolicina_tok,kolicina_do,vrijednost_tok,vrijednost_do,datumunosa,mmgggg) values (" + kkal16 + "," + vkal16 + ", '" + dat1 + "'," + wn1 + ",'" + nkupac16 + "'," + ((int)(k16)).ToString() + "," + ((int)(q16)).ToString() + "," + kupac16 + "," + b16 + "," + d160 + "," + c16.ToString() + "," + e16.ToString() + "," + f16.ToString().Replace(',', '.') + "," + g16.ToString().Replace(',', '.') + ",'" + sada + "','" + mmgggg1 + "');" +
                                "insert into dpr (kolicina_kal,vrijednost_kal,datum,tjedan,kupac,sgr_kom,sgr_eur,id_kupac,planiranokol_tok,planiranokol_do,kolicina_tok,kolicina_do,vrijednost_tok,vrijednost_do,datumunosa,mmgggg) values (" + kkal17 + "," + vkal17 + ", '" + dat1 + "'," + wn1 + ",'" + nkupac17 + "'," + ((int)(k17)).ToString() + "," + ((int)(q17)).ToString() + "," + kupac17 + "," + b17 + "," + d170 + "," + c17.ToString() + "," + e17.ToString() + "," + f17.ToString().Replace(',', '.') + "," + g17.ToString().Replace(',', '.') + ",'" + sada + "','" + mmgggg1 + "');" +
                                "insert into dpr (kolicina_kal,vrijednost_kal,datum,tjedan,kupac,sgr_kom,sgr_eur,id_kupac,planiranokol_tok,planiranokol_do,kolicina_tok,kolicina_do,vrijednost_tok,vrijednost_do,datumunosa,mmgggg) values (" + kkal18 + "," + vkal18 + ", '" + dat1 + "'," + wn1 + ",'" + nkupac18 + "'," + ((int)(k18)).ToString() + "," + ((int)(q18)).ToString() + "," + kupac18 + "," + b18 + "," + d180 + "," + c18.ToString() + "," + e18.ToString() + "," + f18.ToString().Replace(',', '.') + "," + g18.ToString().Replace(',', '.') + ",'" + sada + "','" + mmgggg1 + "');" +
                                "insert into dpr (kolicina_kal,vrijednost_kal,datum,tjedan,kupac,sgr_kom,sgr_eur,id_kupac,planiranokol_tok,planiranokol_do,kolicina_tok,kolicina_do,vrijednost_tok,vrijednost_do,datumunosa,mmgggg) values (" + kkal19 + "," + vkal19 + ", '" + dat1 + "'," + wn1 + ",'" + nkupac19 + "'," + ((int)(k19)).ToString() + "," + ((int)(q19)).ToString() + "," + kupac19 + "," + b19 + "," + d190 + "," + c19.ToString() + "," + e19.ToString() + "," + f19.ToString().Replace(',', '.') + "," + g19.ToString().Replace(',', '.') + ",'" + sada + "','" + mmgggg1 + "');" +
                                "insert into dpr (kolicina_kal,vrijednost_kal,datum,tjedan,kupac,sgr_kom,sgr_eur,id_kupac,planiranokol_tok,planiranokol_do,kolicina_tok,kolicina_do,vrijednost_tok,vrijednost_do,datumunosa,mmgggg) values (" + kkal21 + "," + vkal21 + ", '" + dat1 + "'," + wn1 + ",'" + nkupac21 + "'," + ((int)(k21)).ToString() + "," + ((int)(q21)).ToString() + "," + kupac21 + "," + b21 + "," + d210 + "," + c21.ToString() + "," + e21.ToString() + "," + f21.ToString().Replace(',', '.') + "," + g21.ToString().Replace(',', '.') + ",'" + sada + "','" + mmgggg1 + "');"+
                                "insert into dpr (kolicina_kal,vrijednost_kal,datum,tjedan,kupac,sgr_kom,sgr_eur,id_kupac,planiranokol_tok,planiranokol_do,kolicina_tok,kolicina_do,vrijednost_tok,vrijednost_do,datumunosa,mmgggg) values (" + kkal23 + "," + vkal23 + ", '" + dat1 + "'," + wn1 + ",'" + nkupac23 + "'," + ((int)(k23)).ToString() + "," + ((int)(q23)).ToString() + "," + kupac23 + "," + b23 + "," + d230 + "," + c23.ToString() + "," + e23.ToString() + "," + f23.ToString().Replace(',', '.') + "," + g23.ToString().Replace(',', '.') + ",'" + sada + "','" + mmgggg1 + "');";


                        //sql1 = "insert into dpr  (sgr_kom,sgr_eur,datum,kupac,id_kupac,planiranokol_tok,planiranokol_do,kolicina_tok,kolicina_do,vrijednost_tok,vrijednost_do,kolicina_kal,vrijednost_kal,datumunosa,mmgggg) values (" + kk7 + "," + qq7 + ",'" + dat1 + "','" + nkupac7 + "'," + kupac7 + "," + b7.ToString() + "," + d70.ToString() + "," + c7.ToString() + "," + e7.ToString() + "," + f7.ToString().Replace(',', '.') + "," + g7.ToString().Replace(',', '.') + "," + kkal7 + "," + vkal7 + ",'" + sada + "','" + mmgggg1 + "');" +
                        //                             "insert into dpr  (sgr_kom,sgr_eur,datum,kupac,id_kupac,planiranokol_tok,planiranokol_do,kolicina_tok,kolicina_do,vrijednost_tok,vrijednost_do,kolicina_kal,vrijednost_kal,datumunosa,mmgggg) values (" + kk8 + "," + qq7 + ",'" + dat1 + "','" + nkupac8 + "'," + kupac8 + "," + b8.ToString() + "," + d80.ToString() + "," + c8.ToString() + "," + e8.ToString() + "," + f8.ToString().Replace(',', '.') + "," + g8.ToString().Replace(',', '.') + "," + kkal8 + "," + vkal8 + ",'" + sada + "','" + mmgggg1 + "');" +
                        //                               "insert into dpr (sgr_kom,sgr_eur,datum,kupac,id_kupac,planiranokol_tok,planiranokol_do,kolicina_tok,kolicina_do,vrijednost_tok,vrijednost_do,kolicina_kal,vrijednost_kal,datumunosa,mmgggg) values (" + kk9 + "," + qq9 + ",'" + dat1 + "','" + nkupac9 + "'," + kupac9 + "," + b9.ToString() + "," + d90.ToString() + "," + c9.ToString() + "," + e9.ToString() + "," + f9.ToString().Replace(',', '.') + "," + g9.ToString().Replace(',', '.') + "," + kkal9 + "," + vkal9 + ",'" + sada + "','" + mmgggg1 + "');" +
                        //                               "insert into dpr (sgr_kom,sgr_eur,datum,kupac,id_kupac,planiranokol_tok,planiranokol_do,kolicina_tok,kolicina_do,vrijednost_tok,vrijednost_do,kolicina_kal,vrijednost_kal,datumunosa,mmgggg) values (" + kk10 + "," + qq10 + ",'" + dat1 + "','" + nkupac10 + "'," + kupac10 + "," + b10.ToString() + "," + d90.ToString() + "," + c10.ToString() + "," + e10.ToString() + "," + f10.ToString().Replace(',', '.') + "," + g10.ToString().Replace(',', '.') + "," + kkal10 + "," + vkal10 + ",'" + sada + "','" + mmgggg1 + "');" +
                        //                               "insert into dpr (sgr_kom,sgr_eur,datum,kupac,id_kupac,planiranokol_tok,planiranokol_do,kolicina_tok,kolicina_do,vrijednost_tok,vrijednost_do,kolicina_kal,vrijednost_kal,datumunosa,mmgggg) values (" + kk11 + "," + qq11 + ",'" + dat1 + "','" + nkupac11 + "'," + kupac11 + "," + b11.ToString() + "," + d110.ToString() + "," + c11.ToString() + "," + e11.ToString() + "," + f11.ToString().Replace(',', '.') + "," + g11.ToString().Replace(',', '.') + "," + kkal11 + "," + vkal11 + ",'" + sada + "','" + mmgggg1 + "');" +
                        //                               "insert into dpr (sgr_kom,sgr_eur,datum,kupac,id_kupac,planiranokol_tok,planiranokol_do,kolicina_tok,kolicina_do,vrijednost_tok,vrijednost_do,kolicina_kal,vrijednost_kal,datumunosa,mmgggg) values (" + kk12 + "," + qq12 + ",'" + dat1 + "','" + nkupac12 + "'," + kupac12 + "," + b12.ToString() + "," + d120.ToString() + "," + c12.ToString() + "," + e12.ToString() + "," + f12.ToString().Replace(',', '.') + "," + g12.ToString().Replace(',', '.') + "," + kkal12 + "," + vkal12 + ",'" + sada + "','" + mmgggg1 + "');" +
                        //                               "insert into dpr (sgr_kom,sgr_eur,datum,kupac,id_kupac,planiranokol_tok,planiranokol_do,kolicina_tok,kolicina_do,vrijednost_tok,vrijednost_do,kolicina_kal,vrijednost_kal,datumunosa,mmgggg) values (" + kk13 + "," + qq13 + ",'" + dat1 + "','" + nkupac13 + "'," + kupac13 + "," + b13.ToString() + "," + d130.ToString() + "," + c13.ToString() + "," + e13.ToString() + "," + f13.ToString().Replace(',', '.') + "," + g13.ToString().Replace(',', '.') + "," + kkal13 + "," + vkal13 + ",'" + sada + "','" + mmgggg1 + "');" +
                        //                               "insert into dpr (sgr_kom,sgr_eur,datum,kupac,id_kupac,planiranokol_tok,planiranokol_do,kolicina_tok,kolicina_do,vrijednost_tok,vrijednost_do,kolicina_kal,vrijednost_kal,datumunosa,mmgggg) values (" + kk14 + "," + qq14 + ",'" + dat1 + "','" + nkupac14 + "'," + kupac14 + "," + b14.ToString() + "," + d140.ToString() + "," + c14.ToString() + "," + e14.ToString() + "," + f14.ToString().Replace(',', '.') + "," + g14.ToString().Replace(',', '.') + "," + kkal14 + "," + vkal14 + ",'" + sada + "','" + mmgggg1 + "');" +
                        //                               "insert into dpr (sgr_kom,sgr_eur,datum,kupac,id_kupac,planiranokol_tok,planiranokol_do,kolicina_tok,kolicina_do,vrijednost_tok,vrijednost_do,kolicina_kal,vrijednost_kal,datumunosa,mmgggg) values (" + kk15 + "," + qq15 + ",'" + dat1 + "','" + nkupac15 + "'," + kupac15 + "," + b15.ToString() + "," + d150.ToString() + "," + c15.ToString() + "," + e15.ToString() + "," + f15.ToString().Replace(',', '.') + "," + g15.ToString().Replace(',', '.') + "," + kkal15 + "," + vkal15 + ",'" + sada + "','" + mmgggg1 + "');" +
                        //                               "insert into dpr (sgr_kom,sgr_eur,datum,kupac,id_kupac,planiranokol_tok,planiranokol_do,kolicina_tok,kolicina_do,vrijednost_tok,vrijednost_do,kolicina_kal,vrijednost_kal,datumunosa,mmgggg) values (" + kk16 + "," + qq16 + ",'" + dat1 + "','" + nkupac16 + "'," + kupac16 + "," + b16.ToString() + "," + d160.ToString() + "," + c16.ToString() + "," + e16.ToString() + "," + f16.ToString().Replace(',', '.') + "," + g16.ToString().Replace(',', '.') + "," + kkal16 + "," + vkal16 + ",'" + sada + "','" + mmgggg1 + "');" +
                        //                               "insert into dpr (sgr_kom,sgr_eur,datum,kupac,id_kupac,planiranokol_tok,planiranokol_do,kolicina_tok,kolicina_do,vrijednost_tok,vrijednost_do,kolicina_kal,vrijednost_kal,datumunosa,mmgggg) values (" + kk17 + "," + qq17 + ",'" + dat1 + "','" + nkupac17 + "'," + kupac17 + "," + b17.ToString() + "," + d170.ToString() + "," + c17.ToString() + "," + e17.ToString() + "," + f17.ToString().Replace(',', '.') + "," + g17.ToString().Replace(',', '.') + "," + kkal17 + "," + vkal17 + ",'" + sada + "','" + mmgggg1 + "');" +
                        //                               "insert into dpr (sgr_kom,sgr_eur,datum,kupac,id_kupac,planiranokol_tok,planiranokol_do,kolicina_tok,kolicina_do,vrijednost_tok,vrijednost_do,kolicina_kal,vrijednost_kal,datumunosa,mmgggg) values (" + kk18 + "," + qq18 + ",'" + dat1 + "','" + nkupac18 + "'," + kupac18 + "," + b18.ToString() + "," + d180.ToString() + "," + c18.ToString() + "," + e18.ToString() + "," + f18.ToString().Replace(',', '.') + "," + g18.ToString().Replace(',', '.') + "," + kkal18 + "," + vkal18 + ",'" + sada + "','" + mmgggg1 + "');" +
                        //                               "insert into dpr (sgr_kom,sgr_eur,datum,kupac,id_kupac,planiranokol_tok,planiranokol_do,kolicina_tok,kolicina_do,vrijednost_tok,vrijednost_do,kolicina_kal,vrijednost_kal,datumunosa,mmgggg) values (" + kk19 + "," + qq19 + ",'" + dat1 + "','" + nkupac19 + "'," + kupac19 + "," + b19.ToString() + "," + d190.ToString() + "," + c19.ToString() + "," + e19.ToString() + "," + f19.ToString().Replace(',', '.') + "," + g19.ToString().Replace(',', '.') + "," + kkal19 + "," + vkal19 + ",'" + sada + "','" + mmgggg1 + "');" +
                        //                               "insert into dpr (sgr_kom,sgr_eur,datum,kupac,id_kupac,planiranokol_tok,planiranokol_do,kolicina_tok,kolicina_do,vrijednost_tok,vrijednost_do,kolicina_kal,vrijednost_kal,datumunosa,mmgggg) values (" + kk21 + "," + qq21 + ",'" + dat1 + "','" + nkupac21 + "'," + kupac21 + "," + b21.ToString() + "," + d210.ToString() + "," + c21.ToString() + "," + e21.ToString() + "," + f21.ToString().Replace(',', '.') + "," + g21.ToString().Replace(',', '.') + "," + kkal21 + "," + vkal21 + ",'" + sada + "','" + mmgggg1 + "');";


                        SqlCommand sqlCommand = new SqlCommand(sql1, cn2);
                        SqlDataReader reader = sqlCommand.ExecuteReader();
                        cn2.Close();

                        sql1 = "update dpr set cm = 0, cy = 0, cw = 0;" +
                             "update dpr set cm = 1 where month(datum)= month(getdate()) and year(datum)= year(getdate());" +
                             "update dpr set cm = 2 where month(datum)= month(getdate())-1 and year(datum)= year(getdate());" +
                             "update dpr set cy = 1 where year(datum)= year(getdate());  " +
                             "update dpr set cy = 2 where year(datum)= year(getdate())-1;  " +
                             "update dpr set cw = 1 where year(datum)= year(getdate()) and datepart(week, datum)= datepart(week, dateadd(day, -1, getdate()));" +
                             "update dpr set cw = 2 where year(datum)= year(getdate()) and datepart(week, datum)= datepart(week, dateadd(day, -1, getdate())) - 1";
                        cn2.Open();
                        sqlCommand = new SqlCommand(sql1, cn2);
                        reader = sqlCommand.ExecuteReader();
                        cn2.Close();

                        ///////////////// update planiranjeshare
                        cn2.Open();
                        sql1 = "insert into planiranje.dbo.planiranjeshare(datum, proizvedeno_tok, proizvedeno_do, vrijednost_tok, vrijednost_do, tjedan, kupac, lgr_kom, lgr_eur) select datum, kolicina_tok, kolicina_do, vrijednost_tok, vrijednost_do, tjedan, id_kupac, sgr_kom, sgr_eur from rfind.dbo.dpr where datum='"+dat1+"'";
                        sqlCommand = new SqlCommand(sql1, cn2);
                        reader = sqlCommand.ExecuteReader();
                        cn2.Close();


                    }
                }


                    worksheet.Rows.Cells[7, 13].value = i7;   // broj linija
                worksheet.Rows.Cells[8, 13].value = i8;
                worksheet.Rows.Cells[9, 13].value = i9;
                worksheet.Rows.Cells[10, 13].value = i10;
                worksheet.Rows.Cells[11, 13].value = i11;
                worksheet.Rows.Cells[12, 13].value = i12;
                worksheet.Rows.Cells[13, 13].value = i13;
                worksheet.Rows.Cells[14, 13].value = i14;
                worksheet.Rows.Cells[15, 13].value = i15;
                worksheet.Rows.Cells[16, 13].value = i151;
                worksheet.Rows.Cells[17, 13].value = i16;
                worksheet.Rows.Cells[18, 13].value = i17;
                worksheet.Rows.Cells[19, 13].value = i18;
                worksheet.Rows.Cells[20, 13].value = i19;
                worksheet.Rows.Cells[22, 13].value = i21;
                worksheet.Rows.Cells[23, 13].value = i22 - worksheet.Rows.Cells[25, 13].value;

                worksheet.Rows.Cells[7, 15].value = j7;   // broj štelanja
                worksheet.Rows.Cells[8, 15].value = j8;
                worksheet.Rows.Cells[9, 15].value = j9;
                worksheet.Rows.Cells[10, 15].value = j10;
                worksheet.Rows.Cells[11, 15].value = j11;
                worksheet.Rows.Cells[12, 15].value = j12;
                worksheet.Rows.Cells[13, 15].value = j13;
                worksheet.Rows.Cells[14, 15].value = j14;
                worksheet.Rows.Cells[15, 15].value = j15;
                worksheet.Rows.Cells[16, 15].value = j151;
                worksheet.Rows.Cells[17, 15].value = j16;
                worksheet.Rows.Cells[18, 15].value = j17;
                worksheet.Rows.Cells[19, 15].value = j18;
                worksheet.Rows.Cells[20, 15].value = j19;
                worksheet.Rows.Cells[22, 15].value = j21;

                if (ws <= 3)
                {

                    if (ws == 1)
                    {
                        worksheet.Rows.Cells[7, 16].value = k7;   // LGR, komada
                        worksheet.Rows.Cells[8, 16].value = k8;
                        worksheet.Rows.Cells[9, 16].value = k9;
                        worksheet.Rows.Cells[10, 16].value = k10;
                        worksheet.Rows.Cells[11, 16].value = k11;
                        worksheet.Rows.Cells[12, 16].value = k12;
                        worksheet.Rows.Cells[13, 16].value = k13;
                        worksheet.Rows.Cells[14, 16].value = k14;
                        worksheet.Rows.Cells[15, 16].value = k15;
                        worksheet.Rows.Cells[16, 16].value = k151;
                        worksheet.Rows.Cells[17, 16].value = k16;
                        worksheet.Rows.Cells[18, 16].value = k17;
                        worksheet.Rows.Cells[19, 16].value = k18;
                        worksheet.Rows.Cells[20, 16].value = k19;
                        worksheet.Rows.Cells[22, 16].value = k21;

                        worksheet.Rows.Cells[7, 17].value = q7;   // LGR, vrijednost
                        worksheet.Rows.Cells[8, 17].value = q8;
                        worksheet.Rows.Cells[9, 17].value = q9;
                        worksheet.Rows.Cells[10, 17].value = q10;
                        worksheet.Rows.Cells[11, 17].value = q11;
                        worksheet.Rows.Cells[12, 17].value = q12;
                        worksheet.Rows.Cells[13, 17].value = q13;
                        worksheet.Rows.Cells[14, 17].value = q14;
                        worksheet.Rows.Cells[15, 17].value = q15;
                        worksheet.Rows.Cells[16, 17].value = q151;
                        worksheet.Rows.Cells[17, 17].value = q16;
                        worksheet.Rows.Cells[18, 17].value = q17;
                        worksheet.Rows.Cells[19, 17].value = q18;
                        worksheet.Rows.Cells[20, 17].value = q19;
                        worksheet.Rows.Cells[22, 17].value = q21;

                    }

                }

                worksheet.Rows.Cells[7, 14].value = n7;   // broj linija koje ne rade
                worksheet.Rows.Cells[8, 14].value = n8;
                worksheet.Rows.Cells[9, 14].value = n9;
                worksheet.Rows.Cells[10, 14].value = n10;
                worksheet.Rows.Cells[11, 14].value = n11;
                worksheet.Rows.Cells[12, 14].value = n12;
                worksheet.Rows.Cells[13, 14].value = n13;
                worksheet.Rows.Cells[14, 14].value = n14;
                worksheet.Rows.Cells[15, 14].value = n15;                
                worksheet.Rows.Cells[16, 14].value = n151;
                worksheet.Rows.Cells[17, 14].value = n16;
                worksheet.Rows.Cells[18, 14].value = n17;
                worksheet.Rows.Cells[19, 14].value = n18;
                worksheet.Rows.Cells[20, 14].value = n19;
                worksheet.Rows.Cells[22, 14].value = n21;
                worksheet.Rows.Cells[23, 14].value = worksheet.Rows.Cells[23, 13].value;    // linija nije  zaplanirana


                // škart postoci
                worksheet.Rows.Cells[7, 9].value = Math.Round(post7t, 4);   // škart obrade
                worksheet.Rows.Cells[8, 9].value = Math.Round(post8t, 4);
                worksheet.Rows.Cells[9, 9].value = Math.Round(post9t, 4);
                worksheet.Rows.Cells[10, 9].value = Math.Round(post10t, 4);
                worksheet.Rows.Cells[11, 9].value = Math.Round(post11t, 4);
                worksheet.Rows.Cells[12, 9].value = Math.Round(post12t, 4);
                double posto13 = Math.Round(post13t, 4);
                worksheet.Rows.Cells[13, 9].value = posto13;
                worksheet.Rows.Cells[14, 9].value = Math.Round(post14t, 4);
                worksheet.Rows.Cells[15, 9].value = Math.Round(post15t, 4);
                worksheet.Rows.Cells[16, 9].value = Math.Round(post151t, 4);
                worksheet.Rows.Cells[17, 9].value = Math.Round(post16t, 4);
                worksheet.Rows.Cells[18, 9].value = Math.Round(post17t, 4);
                worksheet.Rows.Cells[19, 9].value = Math.Round(post18t, 4);
                worksheet.Rows.Cells[20, 9].value = Math.Round(post19t, 4);
                worksheet.Rows.Cells[22, 9].value = Math.Round(post21t, 4);
                double pos131 = Math.Round(post13tm, 4);


                worksheet.Rows.Cells[7, 11].value = Math.Round(post7tm, 4);   // škart materijala
                worksheet.Rows.Cells[8, 11].value = Math.Round(post8tm, 4);
                worksheet.Rows.Cells[9, 11].value = Math.Round(post9tm, 4);
                worksheet.Rows.Cells[9, 11].value = Math.Round(post10tm, 4);
                worksheet.Rows.Cells[11, 11].value = Math.Round(post11tm, 4);
                worksheet.Rows.Cells[12, 11].value = Math.Round(post12tm, 4);
                worksheet.Rows.Cells[13, 11].value = Math.Round(post13tm, 4);
                worksheet.Rows.Cells[14, 11].value = Math.Round(post14tm, 4);
                worksheet.Rows.Cells[15, 11].value = Math.Round(post15tm, 4);
                worksheet.Rows.Cells[16, 11].value = Math.Round(post151tm, 4);
                worksheet.Rows.Cells[17, 11].value = Math.Round(post16tm, 4);
                worksheet.Rows.Cells[18, 11].value = Math.Round(post17tm, 4);
                worksheet.Rows.Cells[19, 11].value = Math.Round(post18tm, 4);
                worksheet.Rows.Cells[20, 11].value = Math.Round(post19tm, 4);
                worksheet.Rows.Cells[22, 11].value = Math.Round(post21tm, 4);

                // škart komadi
                worksheet.Rows.Cells[7, 10].value = e7to;   // škart obrade
                worksheet.Rows.Cells[8, 10].value = e8to;
                worksheet.Rows.Cells[9, 10].value = e9to;
                worksheet.Rows.Cells[10, 10].value = e10to;
                worksheet.Rows.Cells[11, 10].value = e11to;
                worksheet.Rows.Cells[12, 10].value = e12to;

                worksheet.Rows.Cells[13, 10].value = e13to;
                worksheet.Rows.Cells[14, 10].value = e14to;
                worksheet.Rows.Cells[15, 10].value = e15to;
                worksheet.Rows.Cells[16, 10].value = e151to;
                worksheet.Rows.Cells[17, 10].value = e16to;
                worksheet.Rows.Cells[18, 10].value = e16to;
                worksheet.Rows.Cells[19, 10].value = e18to;
                worksheet.Rows.Cells[20, 10].value = e19to;
                worksheet.Rows.Cells[22, 10].value = e21to;

                worksheet.Rows.Cells[7, 12].value = e7ts;   // škart materijala
                worksheet.Rows.Cells[8, 12].value = e8ts;
                worksheet.Rows.Cells[9, 12].value = e9ts;
                worksheet.Rows.Cells[10, 12].value = e10ts;
                worksheet.Rows.Cells[11, 12].value = e11ts;
                worksheet.Rows.Cells[12, 12].value = e12ts;
                worksheet.Rows.Cells[13, 12].value = e13ts;
                worksheet.Rows.Cells[14, 12].value = e14ts;
                worksheet.Rows.Cells[15, 12].value = e15ts;
                worksheet.Rows.Cells[16, 12].value = e151ts;
                worksheet.Rows.Cells[17, 12].value = e16ts;
                worksheet.Rows.Cells[18, 12].value = e17ts;
                worksheet.Rows.Cells[19, 10].value = e18ts;
                worksheet.Rows.Cells[20, 12].value = e19ts;
                worksheet.Rows.Cells[22, 12].value = e21ts;

                
                //            worksheet.Rows.Cells[26, 2].value = b26;    // Realizacija ukupno: planirano
                worksheet.Rows.Cells[29, 2].value = worksheet.Rows.Cells[25, 2].value + worksheet.Rows.Cells[25, 4].value;
                worksheet.Rows.Cells[29, 3].value = worksheet.Rows.Cells[25, 3].value + worksheet.Rows.Cells[25, 5].value;     // ukupna količina

                //           worksheet.Rows.Cells[27, 2].value = b27;    // Realizacija ukupno: realizirano
                worksheet.Rows.Cells[30, 3].value = worksheet.Rows.Cells[25, 6].value + worksheet.Rows.Cells[25, 7].value;     // ukupna količinac27;    // ukupna vrijednost
                

                if (ws == 2 || ws == 3)
                {
                    worksheet.Rows.Cells[34, 2].value = b312;    // kaliona
                    worksheet.Rows.Cells[34, 5].value = e312;    // kaliona
                    worksheet.Rows.Cells[34, 9].value = (e312 + b312) * 0.33;  // in
                    worksheet.Rows.Cells[34, 7].value = g31;   // broj šarži
                    worksheet.Rows.Cells[34, 10].value = (e312 + b312);   // in   // broj šarži
                }


                if (ws == 1)
                {
                    worksheet.Rows.Cells[34, 3].value = b31;    // kaliona
                    worksheet.Rows.Cells[34, 6].value = e31;
                    worksheet.Rows.Cells[34, 8].value = g31;   // broj šarži
                    worksheet.Rows.Cells[34, 10].value = (e31 + b31);   // in
                    worksheet.Rows.Cells[34, 11].value = (e31 + b31) * 0.33;  // in
                    worksheet.Rows.Cells[35, 3].value = b312;    // kaliona
                    worksheet.Rows.Cells[35, 6].value = e312;
                    //                worksheet.Rows.Cells[34, 7].value = g31;   // broj šarži
                    worksheet.Rows.Cells[35, 10].value = (e312 + b312);
                    worksheet.Rows.Cells[35, 11].value = (e312 + b312) * 0.33;  // out
                    worksheet.Rows.Cells[2, 15].value = worksheet.Rows.Cells[2, 15].value + " " + dat10;
                    //  worksheet.Rows.Cells[2, 20].value = worksheet.Rows.Cells[2, 20].value + " " + dat10;
                    //  worksheet.Rows.Cells[9, 20].value = dat10;

                    worksheet.Rows.Cells[42, 1].value = dat30;

                    //worksheet.Rows.Cells[36, 1].value = worksheet.Rows.Cells[34, 1].value + " " + dat10;
                    worksheet.Rows.Cells[29, 6].value = ul24;    // Transporti 24 t
                    worksheet.Rows.Cells[29, 7].value = iz24;
                    worksheet.Rows.Cells[30, 6].value = ul15;    // Transporti 1.5T
                    worksheet.Rows.Cells[30, 7].value = iz15;

                    worksheet.Rows.Cells[43, 2].value = b39;   // količina materijala, HAY
                    worksheet.Rows.Cells[43, 3].value = c39;   // količina materijala
                    worksheet.Rows.Cells[43, 4].value = d39;   // količina materijala
                    worksheet.Rows.Cells[43, 5].value = e39;   // količina materijala
                    worksheet.Rows.Cells[43, 6].value = f39;   // količina materijala
                    worksheet.Rows.Cells[43, 7].value = g39;   // količina materijala
                    worksheet.Rows.Cells[43, 8].value = h39;   // količina materijala
                    worksheet.Rows.Cells[43, 9].value = i39;   // količina materijala
                    worksheet.Rows.Cells[43, 10].value = j39;   // količina materijala
                    worksheet.Rows.Cells[43, 11].value = k39;   // količina materijala
                    worksheet.Rows.Cells[43, 12].value = l39;   // količina materijala
                    worksheet.Rows.Cells[43, 13].value = m39;   // količina materijala
                    worksheet.Rows.Cells[43, 14].value = n39;   // količina materijala
                    worksheet.Rows.Cells[43, 15].value = o39;   // količina materijala
                    worksheet.Rows.Cells[43, 16].value = p39;   // količina materijala

                    worksheet.Rows.Cells[44, 2].value = b42;   // vrijednost materijala, HAY
                    worksheet.Rows.Cells[44, 3].value = c42;   // vrijednost materijala
                    worksheet.Rows.Cells[44, 4].value = d42;   // vrijednost materijala
                    worksheet.Rows.Cells[44, 5].value = e42;   // vrijednost materijala
                    worksheet.Rows.Cells[44, 6].value = f42;   // vrijednost materijala
                    worksheet.Rows.Cells[44, 7].value = g42;   // vrijednost materijala
                    worksheet.Rows.Cells[44, 8].value = h42;   // vrijednost materijala
                    worksheet.Rows.Cells[44, 9].value = i42;   // vrijednost materijala
                    worksheet.Rows.Cells[44, 10].value = j42;   // vrijednost materijala
                    worksheet.Rows.Cells[44, 11].value = k42;   // vrijednost materijala
                    worksheet.Rows.Cells[44, 12].value = l42;   // vrijednost materijala
                    worksheet.Rows.Cells[44, 13].value = m42;   // vrijednost materijala
                    worksheet.Rows.Cells[44, 14].value = n42;   // vrijednost materijala
                    worksheet.Rows.Cells[44, 15].value = o42;   // vrijednost materijala
                    worksheet.Rows.Cells[44, 16].value = p42;   // vrijednost materijala

                    //worksheet.Rows.Cells[7, 20].value = naBolovanju;   // Izostanci, na bolovanju
                    //worksheet.Rows.Cells[7, 21].value = naGO;   // na godišnjem
                    //worksheet.Rows.Cells[7, 22].value = ostalo;   // ostalo
                    //worksheet.Rows.Cells[8, 22].value = ostalo+naGO+naBolovanju ;   // ukupno
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
            Worksheet worksheet3r2 = workbook.Worksheets.Item[6] as Worksheet;   // ldp po radnicima
            Worksheet worksheet4l = workbook.Worksheets.Item[7] as Worksheet;    // ldp po linijama

            Worksheet workshstelanje = workbook.Worksheets.Item[8] as Worksheet;    // Pregled štelanja,aktivnosti
            Worksheet worksizostanci = workbook.Worksheets.Item[9] as Worksheet;    // Pregled izostanaka
            Worksheet worksinorme = workbook.Worksheets.Item[10] as Worksheet;    // Pregled normi
            //Worksheet worksrealizacija = workbook.Worksheets.Item[10] as Worksheet;    // Pregled realizacija za pločice


            // LDP
            // radnici

            worksheet3r2.Rows.Cells[1, 2].value = worksheet3r2.Rows.Cells[1, 2].value + " " + datLDPTitle;
            //            worksheet3r.Rows.Cells[2, 6].value = smjenanaziv;
            // linije
            worksheet4l.Rows.Cells[1, 2].value = worksheet4l.Rows.Cells[1, 2].value + " " + datLDPTitle;
            //          worksheet4l.Rows.Cells[2, 6].value = smjenanaziv;


            // pregled izostanka, popis imena,mt

            using (SqlConnection cn = new SqlConnection(connectionString2))
            {
                cn.Open();
                //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                sql1 = "rfind.dbo.ldp_recalc2 '" + datLDP + "','" + datLDP + "',71";
                SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                SqlDataReader reader = sqlCommand.ExecuteReader();
                worksizostanci.Rows.Cells[4, 3].value = worksizostanci.Rows.Cells[4, 3].value + " " + datreps;
                string vrstao = "";
                int i = 0;
                while (reader.Read())
                {
                    if (1 == 1)
                    {

                        worksizostanci.Rows.Cells[8 + i, 1].value = reader["mtroska"];
                        worksizostanci.Rows.Cells[8 + i, 2].value = reader["hala"];
                        worksizostanci.Rows.Cells[8 + i, 3].value = reader["smjena"];
                        worksizostanci.Rows.Cells[8 + i, 4].value = reader["linija"].ToString();
                        worksizostanci.Rows.Cells[8 + i, 5].value = reader["prezime"].ToString();
                        worksizostanci.Rows.Cells[8 + i, 6].value = reader["ime"].ToString();
                        worksizostanci.Rows.Cells[8 + i, 7].value = reader["vrsta"].ToString();


                    }
                    i++;

                }

                cn.Close();
            }
            Console.WriteLine("Izostanci  trenutno vrijeme " + DateTime.Now);

            // pregled normi=1

            using (SqlConnection cn = new SqlConnection(connectionString2))
            {
                cn.Open();
                sql1 = "rfind.dbo.ldp_recalc '" + datLDP + "','" + datLDP + "',10";
                SqlCommand sqlCommand = new SqlCommand(sql1, cn);
                SqlDataReader reader = sqlCommand.ExecuteReader();

                //worksinorme.Rows.Cells[4, 3].value = worksinorme.Rows.Cells[4, 3].value + " " + datreps;               
                int i =3;

                while (reader.Read())
                {
                    if (1 == 1)
                    {
                        worksinorme.Rows.Cells[2 + i, 1].value = reader["datum"];
                        worksinorme.Rows.Cells[2 + i, 2].value = reader["hala"];
                        worksinorme.Rows.Cells[2 + i, 3].value = reader["linija"].ToString();
                        worksinorme.Rows.Cells[2 + i, 4].value = reader["id_pro"].ToString();
                        worksinorme.Rows.Cells[2 + i, 5].value = reader["id_mat"].ToString();
                        worksinorme.Rows.Cells[2 + i, 6].value = reader["nazivpro"].ToString();
                        worksinorme.Rows.Cells[2 + i, 7].value = reader["kolicinaok"].ToString();
                        worksinorme.Rows.Cells[2 + i, 8].value = reader["kupac"].ToString();
                    }
                    i++;

                }

                cn.Close();
            }
            Console.WriteLine("Pregled normi=1  trenutno vrijeme " + DateTime.Now);


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

                        //    worksheet3r.Rows.Cells[5 + i, 1].value = reader["datum"];
                        //    worksheet3r.Rows.Cells[5 + i, 2].value = (reader["hala"].ToString());
                        //    worksheet3r.Rows.Cells[5 + i, 3].value = (reader["linija"].ToString());
                        //    worksheet3r.Rows.Cells[5 + i, 4].value = (reader["nazivpro"].ToString());
                        //    worksheet3r.Rows.Cells[5 + i, 5].value = (reader["radnik"].ToString());
                        //    worksheet3r.Rows.Cells[5 + i, 6].value = reader["norma"];
                        //    worksheet3r.Rows.Cells[5 + i, 7].value = reader["planirano"];
                        //    worksheet3r.Rows.Cells[5 + i, 8].value = (reader["kolicina"].ToString());
                        //    worksheet3r.Rows.Cells[5 + i, 9].value = (reader["kupac"].ToString());
                        //    worksheet3r.Rows.Cells[5 + i, 10].value = (reader["napomena"].ToString());
                        //    worksheet3r.Rows.Cells[5 + i, 11].value = (reader["aktivnost"].ToString());

                        //    if ((int.Parse)(reader["kolicina"].ToString()) > (int.Parse)(reader["planirano"].ToString()))
                        //        worksheet3r.Rows.Cells[5 + i, 8].Interior.Color = XlRgbColor.rgbGreenYellow;

                        //    ukupn = ukupn + (int.Parse)(reader["norma"].ToString());

                        //    if (reader["planirano"] != DBNull.Value)
                        //    {
                        //        ukupp = ukupp + (int.Parse)(reader["planirano"].ToString());
                        //        ukupkol = ukupkol + (int.Parse)(reader["kolicina"].ToString());
                        //    }

                        //    i++;

                    }
                    //worksheet3r.Rows.Cells[5 + i, 5].value = "Ukupno :";
                    //worksheet3r.Rows.Cells[5 + i, 6].value = ukupn;
                    //worksheet3r.Rows.Cells[5 + i, 7].value = ukupp;
                    //worksheet3r.Rows.Cells[5 + i, 8].value = ukupkol;
                    //worksheet3r.Rows.Cells[6 + i, 7].value = "Postotak realizacije";

                    //worksheet3r.Rows.Cells[6 + i, 8].value = Math.Round(((ukupkol + 0.001) / (ukupp + 0.001) * 100), 2);
                }
                else
                {// drugi ldp po radncima

                    if (1 == 1)
                    {
                        sql1 = "select * from rfind.dbo.evidnormiradad('" + datLDP + "','" + datLDP + "') order by vrsta1,vrsta,radnik,hala,smjena,linija"; //" + smj; bez računanja efektivnog radnog vremena  

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

                            worksheet3r2.Rows.Cells[5 + i, 1].value = radnik;
                            worksheet3r2.Rows.Cells[5 + i, 2].value = (reader["vrsta"].ToString());
                            worksheet3r2.Rows.Cells[5 + i, 3].value = (reader["hala"].ToString());
                            worksheet3r2.Rows.Cells[5 + i, 4].value = (reader["smjena"].ToString());
                            worksheet3r2.Rows.Cells[5 + i, 5].value = (reader["linija"].ToString());

                            worksheet3r2.Rows.Cells[5 + i, 6].value = reader["nazivpar"];
                            worksheet3r2.Rows.Cells[5 + i, 7].value = reader["brojrn"];
                            worksheet3r2.Rows.Cells[5 + i, 8].value = (reader["proizvod"].ToString());
                            worksheet3r2.Rows.Cells[5 + i, 9].value = (reader["norma"].ToString());

                            worksheet3r2.Rows.Cells[5 + i, 11].value = (reader["kolicinaok"].ToString());
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
                            worksheet3r2.Rows.Cells[5 + i, 13].value = (reader["otpadobrada"].ToString());
                            worksheet3r2.Rows.Cells[5 + i, 14].value = (reader["kolicinaporozno"].ToString());
                            worksheet3r2.Rows.Cells[5 + i, 15].value = (reader["otpadmat"].ToString());

                            worksheet3r2.Rows.Cells[5 + i, 16].value = reader["minutaradaradnika"];
                            worksheet3r2.Rows.Cells[5 + i, 19].value = (reader["napomena1"].ToString());
                            worksheet3r2.Rows.Cells[5 + i, 20].value = (reader["napomena2"].ToString());
                            worksheet3r2.Rows.Cells[5 + i, 21].value = (reader["napomena3"].ToString());
                            //    double planirano = (satirada * 60.0 - zastoj) * norma / (satirada*60);
                            // begin aktivnosti

                            int trajanje11 = 0;
                            string smjena2= (reader["smjena"].ToString());
                            string hala2 = (reader["hala"].ToString());
                            string linija2 = (reader["linija"].ToString());
                            string radninalog2 = (reader["brojrn"].ToString());

                            using (SqlConnection cna = new SqlConnection(connectionString))
                            {
                                cna.Open();
                                //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                                // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                                dat13 = datLDP + " 06:00:00";
                                dat23 = dat3 + " 5:59:00";

                                if (smjena2 == "1")
                                {
                                    dat13 = datLDP + " 06:00:00";
                                    dat23 = datLDP + " 13:59:00";
                                }
                                if (smjena2 == "2")
                                {
                                    dat13 = datLDP + " 14:00:00";
                                    dat23 = datLDP + " 21:59:00";
                                }
                                if (smjena2 == "3")
                                {
                                    dat13 = datLDP + " 22:00:00";
                                    dat23 = dat3 + " 5:59:00";
                                }

                                string sqla = "select * from rfind.dbo.pregled_po_liniji( '" + dat13 + "','" + dat23 + "','" + hala2 + "','" + linija2.TrimEnd() + "')";
                                SqlCommand sqlCommanda = new SqlCommand(sqla, cna);
                                SqlDataReader readera = sqlCommanda.ExecuteReader();
                                string aktivnost1 = "";
                                string trajanje1 = "";
                                string napomena1 = "", rna1 = "";


                                while (readera.Read())
                                {
                                    if (linija2 == "16")
                                    {
                                        int tt211 = 1;
                                    }

                                    aktivnost1 = readera["aktivnost"].ToString();
                                    trajanje1 = readera["trajanje"].ToString();
                                    if (trajanje1.TrimEnd() == "")
                                        trajanje1 = "0";
                                    //trajanje11 = trajanje11 + (int.Parse)(trajanje1);

                                    napomena1 = readera["napomena"].ToString();
                                    rna1 = readera["brojrn"].ToString().TrimEnd();
                                    if (rna1 != "")
                                    {
                                        if (rna1 != radninalog2)
                                        {
                                            continue;
                                        }
                                    }

                                    if (aktivnost1.Contains("Otežan"))   // Otežan rad
                                    {
                                        worksheet3r2.Rows.Cells[5 + i, 22].value = trajanje1;
                                        worksheet3r2.Rows.Cells[5 + i, 23].value = napomena1 + " ( RN: " + rna1 + " )";
                                    }
                                    if (aktivnost1.Contains("djelatnika"))  // Nedostatak djelatnika
                                    {
                                        worksheet3r2.Rows.Cells[5 + i, 24].value = trajanje1;
                                        worksheet3r2.Rows.Cells[5 + i, 25].value = napomena1 + " ( RN: " + rna1 + " )";

                                    }
                                    if (aktivnost1.Contains("sirovca"))    // Nedostatak sirovca
                                    {
                                        worksheet3r2.Rows.Cells[5 + i, 26].value = trajanje1;
                                        worksheet3r2.Rows.Cells[5 + i, 27].value = napomena1 + " ( RN: " + rna1 + " )";

                                    }
                                    if (aktivnost1.Contains("Dorada"))    // Dorada
                                    {
                                        worksheet3r2.Rows.Cells[5 + i, 28].value = trajanje1;
                                        worksheet3r2.Rows.Cells[5 + i, 29].value = napomena1 + " ( RN: " + rna1 + " )";

                                    }

                                }
                            }
                            
                            // end aktivnosti
                            
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

                                    worksheet3r2.Rows.Cells[5 + i, 12].Interior.Color = XlRgbColor.rgbGreenYellow;
                                }
                                if (posto1 < 0.900)
                                {

                                    worksheet3r2.Rows.Cells[5 + i, 12].Interior.Color = XlRgbColor.rgbRed;
                                }
                                if (posto1 >= 0.900 && posto1 < 1.00)
                                {

                                    worksheet3r2.Rows.Cells[5 + i, 12].Interior.Color = XlRgbColor.rgbOrange;
                                }
                            }

                            worksheet3r2.Rows.Cells[5 + i, 10].value = (int)planirano;
                            worksheet3r2.Rows.Cells[5 + i, 12].value = posto1;


                            ukupn = ukupn + n1;

                            if (reader["norma"] != DBNull.Value)
                            {
                                ukupp = ukupp + n1;
                                ukupkol = ukupkol + k1;
                            }
                            worksheet3r2.Rows.Cells[5 + i, 30].value = worksheet3r2.Rows.Cells[5 + i, 11].value - worksheet3r2.Rows.Cells[5 + i, 9].value;
                            worksheet3r2.Rows.Cells[5 + i, 31].value = worksheet3r2.Rows.Cells[5 + i, 11].value - worksheet3r2.Rows.Cells[5 + i, 10].value;

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
                        worksheet3r2.Rows.Cells[5 + i, 9].value = "Ukupno :";
                        worksheet3r2.Rows.Cells[5 + i, 10].value = ukupn;
                        worksheet3r2.Rows.Cells[5 + i, 11].value = ukupkol;
                        worksheet3r2.Rows.Cells[6 + i, 9].value = "Postotak realizacije";

                        worksheet3r2.Rows.Cells[6 + i, 10].value = Math.Round(((ukupkol + 0.001) / (ukupp + 0.001) * 100), 2);
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
                string hala1 = "", smjena1 = "", proiz1 = "", linija1 = "", radninalog1 = "", radninalog2 = "", napomena10 ="",napomena20="";
                string hala2 = "", smjena2 = "", proiz2 = "", linija2 = "", tt2 = "", tt1 = "";
                string kupac1 = "", kupac2 = "", naziv1 = "" ;
                int norma1 = 0, kolicina1 = 0, kolicina2 = 0, norma2 = 0, norma = 0, kolicina = 0;
                int prvii = 1, normaukup1 = 0, normaukup2 = 0, erv1 = 0, erv2 = 0, erv = 0, normaukupp = 0 , zadnjired=0;
                double cijena1 = 0.0, cijena2 = 0.0, norm = 0.0, norm1 = 0.0;
                bool joss = true;
                string brisi1 = "";                

                while (reader.Read())
                {                    

                    var loop = true;

                    while (loop || (zadnjired == 1))
                    {                          
                       
                        if (!reader.HasRows)
                            {
                            //  loop = false;
                            //continue; //kad nema podataka
                        }

                        try
                        {
                            if (reader["halal"] != DBNull.Value)
                            {
                                //  loop = false;
                                //continue; //kad nema podataka
                            }
                        }
                        catch
                        {
                            loop = false;
                            zadnjired = 0;
                            continue; //kad nema podataka

                        }
                                               
                        if (zadnjired == 0)
                        {
                            if (reader["norma"] != DBNull.Value)
                            {
                                hala1 = (reader["hala"].ToString());
                                if (hala1 == "")
                                {
                                    hala1 = (reader["halal"].ToString());
                                    naziv1 = (reader["naziv"].ToString());
                                    worksheet4l.Rows.Cells[5 + i, 1].value = dat1;
                                    worksheet4l.Rows.Cells[5 + i, 2].value = hala1;
                                    worksheet4l.Rows.Cells[5 + i, 3].value = naziv1;
                                    i++;
                                    reader.Read();
                                    continue;
                                }

                            }
                            else
                            {
                                hala1 = (reader["halal"].ToString());
                                naziv1 = (reader["naziv"].ToString());
                                worksheet4l.Rows.Cells[5 + i, 1].value = dat1;
                                worksheet4l.Rows.Cells[5 + i, 2].value = hala1;
                                worksheet4l.Rows.Cells[5 + i, 3].value = naziv1;
                                i++;
                                reader.Read();
                                continue;
                            }


                            smjena1 = (reader["smjena"].ToString());
                            proiz1 = (reader["nazivpro"].ToString());
                            linija1 = (reader["linija"].ToString());
                            tt1 = (reader["tt"].ToString());
                            radninalog1 = (reader["brojrn"].ToString());
                            kolicina1 = (int.Parse)(reader["kolicinaok"].ToString());
                            normaukup1 = (int.Parse)(reader["normukup"].ToString());
                            erv1 = (int.Parse)(reader["erv"].ToString());
                            norma1 = (int.Parse)(reader["norma"].ToString());
                            cijena1 = (double.Parse)(reader["cijena"].ToString());
                            kupac1 = (reader["kupac"].ToString());
                            norm1 = norma1 * erv1 / normaukup1;
                            napomena10 = (reader["napomena"].ToString());
                            worksheet4l.Rows.Cells[5 + i, 1].value = reader["datum"];
                            brisi1 = (reader["erv"].ToString());

                            if (radninalog1.Contains("NSK"))
                            {
                                erv1 = erv1;
                            }

                            loop = reader.Read();
                        }
                        else
                        {
                            worksheet4l.Rows.Cells[5 + i, 1].value = dat1;
                            hala1 = "X";                            
                        }

                        if ((hala1 == hala2 && linija1 == linija2 && proiz1 == proiz2) || (prvii == 1))
                        {
                            norma = norma + norma1;
                            kolicina = kolicina + kolicina1;
                            erv = erv + erv1;
                            normaukupp = normaukupp + normaukup1;
                            norm = norm + norma1 * erv1 / normaukup1;

                            if (hala1 == hala2 && smjena1 == smjena2 && proiz1 == proiz2 && linija1 == linija2 && radninalog1 != radninalog2)
                            {
                                norma = norma - norma1;
                                // norm  = norm  - norma1 * erv1 / normaukup1;
                                // erv   = erv   - erv1;
                                normaukupp = normaukupp - normaukup1;
                            }
                            //norm = norma * erv / normaukupp; //do 16.7
                            norm = norma;
                            norma1 = norma;
                            erv1 = erv;
                            normaukup1 = normaukupp;
                            norm1 = norm;

                            hala2 = hala1;
                            erv2 = erv1;
                            if (linija2 == "03")
                            {
                                int tt211 = 1;
                            }
                            smjena2 = smjena1;
                            napomena20 = napomena10;
                            proiz2 = proiz1;
                            radninalog2 = radninalog1;
                            linija2 = linija1;
                            tt2 = tt1;
                            cijena2 = cijena1;
                            kupac2 = kupac1;
                            prvii = 0;

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
                        int trajanje11 = 0;
                        using (SqlConnection cna = new SqlConnection(connectionString))
                        {
                            cna.Open();
                            //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                            // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                            dat13 = datLDP + " 06:00:00";
                            dat23 = dat3 + " 5:59:00";

                            if (smjena2 == "1")
                            {
                                dat13 = datLDP + " 06:00:00";
                                dat23 = datLDP + " 13:59:00";
                            }
                            if (smjena2 == "2")
                            {
                                dat13 = datLDP + " 14:00:00";
                                dat23 = datLDP + " 21:59:00";
                            }
                            if (smjena2 == "3")
                            {
                                dat13 = datLDP + " 22:00:00";
                                dat23 = dat3 + " 5:59:00";
                            }

                            string sqla = "select * from rfind.dbo.pregled_po_liniji( '" + dat13 + "','" + dat23 + "','" + hala2 + "','" + linija2.TrimEnd() + "')";
                            SqlCommand sqlCommanda = new SqlCommand(sqla, cna);
                            SqlDataReader readera = sqlCommanda.ExecuteReader();
                            string aktivnost1 = "";
                            string trajanje1 = "";
                            string napomena1 = "", rna1 = "";
                            

                            while (readera.Read())
                            {
                                if (linija2 == "16")
                                {
                                    int tt211 = 1;
                                }

                                aktivnost1 = readera["aktivnost"].ToString();
                                trajanje1 = readera["trajanje"].ToString();
                                if (trajanje1.TrimEnd() == "")
                                        trajanje1 = "0";
                                //trajanje11 = trajanje11 + (int.Parse)(trajanje1);
                                
                                napomena1 = readera["napomena"].ToString();
                                rna1 = readera["brojrn"].ToString().TrimEnd();
                                if (rna1 != "")
                                {
                                    if (rna1 != radninalog2)
                                    {
                                        continue;
                                    }
                                }

                                if (aktivnost1.Contains("telanje"))   // štelanje
                                {
                                    worksheet4l.Rows.Cells[5 + i, 16].value = trajanje1;
                                    worksheet4l.Rows.Cells[5 + i, 17].value = napomena1 + " ( RN: " + rna1 + " )";
                                    trajanje11 = trajanje11 + (int.Parse)(trajanje1);
                                }
                                if (aktivnost1.Contains("Otežan"))   // Otežan rad
                                {
                                    worksheet4l.Rows.Cells[5 + i, 18].value = trajanje1;
                                    worksheet4l.Rows.Cells[5 + i, 19].value = napomena1 + " ( RN: " + rna1 + " )";
                                }
                                if (aktivnost1.Contains("Kvar"))    // Kvar stroja
                                {
                                    worksheet4l.Rows.Cells[5 + i, 20].value = trajanje1;
                                    worksheet4l.Rows.Cells[5 + i, 21].value = napomena1 + " ( RN: " + rna1 + " )";
                                    trajanje11 = trajanje11 + (int.Parse)(trajanje1);
                                }
                                if (aktivnost1.Contains("djelatnika"))  // Nedostatak djelatnika
                                {
                                    worksheet4l.Rows.Cells[5 + i, 22].value = trajanje1;
                                    worksheet4l.Rows.Cells[5 + i, 23].value = napomena1 + " ( RN: " + rna1 + " )";
                                    
                                }
                                if (aktivnost1.Contains("sirovca"))    // Nedostatak sirovca
                                {
                                    worksheet4l.Rows.Cells[5 + i, 20].value = trajanje1;
                                    worksheet4l.Rows.Cells[5 + i, 21].value = napomena1 + " ( RN: " + rna1 + " )";
                                    
                                }
                                if (aktivnost1.Contains("Dorada"))    // Dorada
                                {
                                    worksheet4l.Rows.Cells[5 + i, 22].value = trajanje1;
                                    worksheet4l.Rows.Cells[5 + i, 23].value = napomena1 + " ( RN: " + rna1 + " )";
                                    
                                }

                            }
                        }

                        worksheet4l.Rows.Cells[5 + i, 2].value = hala2;
                        worksheet4l.Rows.Cells[5 + i, 3].value = linija2;
                        if (linija2.Contains("VC510") || linija2.Contains("HURCO") || linija2.Contains("INDEX") || linija2.Contains("GT600") || tt2.Contains("1"))
                        {
                            worksheet4l.Rows.Cells[5 + i, 4].value = "Dodatne operacije";
                        }
                        else
                        {
                            worksheet4l.Rows.Cells[5 + i, 4].value = "Tokarenje";
                        }

                        worksheet4l.Rows.Cells[5 + i, 5].value = proiz2;
                        worksheet4l.Rows.Cells[5 + i, 6].value = norma; //  reader["norma"]; ;  // norma
                        worksheet4l.Rows.Cells[5 + i, 7].value = kolicina; // reader["kolicinaok"]; ;  // komada                    

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
                        norm = norma * (normaukupp - trajanje11) / normaukupp ; // novo poslije 16.7.
                        n1   = norm * cijena2                              ;
                        if (linija2=="05")
                        {
                            n1 = n1;
                        }
                        worksheet4l.Rows.Cells[5 + i, 8].value = norm; // planirano komada do 16.7.
                        //worksheet4l.Rows.Cells[5 + i, 8].value = norm * ( normaukup - trajanje11)/normaukup ; // planirano komada


                        worksheet4l.Rows.Cells[5 + i, 9].value = cijena2;

                        posto1 = (k1 + 0.0000001) / (n1 + 0.0000001);
                        posto1 = (kolicina + 0.0000001) / (norm + 0.0000001);

                        if (kolicina <= 0.0001 || norm <= 0.0001)
                        {
                            posto1 = 0.0;
                        }

                        if (posto1 >= 1.00)
                        {
                            worksheet4l.Rows.Cells[5 + i, 12].Interior.Color = XlRgbColor.rgbGreenYellow;
                        }
                        if (posto1 < 0.900)
                        {
                            worksheet4l.Rows.Cells[5 + i, 12].Interior.Color = XlRgbColor.rgbRed;
                        }
                        if (posto1 >= 0.900 && posto1 < 1.00)
                        {
                            worksheet4l.Rows.Cells[5 + i, 12].Interior.Color = XlRgbColor.rgbOrange;
                        }

                        worksheet4l.Rows.Cells[5 + i, 10].value = ostvareno;
                        worksheet4l.Rows.Cells[5 + i, 11].value = n1;   // korigirana norma za srtvarno vrijeme rada
                        worksheet4l.Rows.Cells[5 + i, 12].value = posto1;
                        worksheet4l.Rows.Cells[5 + i, 13].value = ostvareno - planirano;
                        //   worksheet4l.Rows.Cells[5 + i, 12].value = (reader["napomena"].ToString());
                        worksheet4l.Rows.Cells[5 + i, 15].value = kupac2;
                        worksheet4l.Rows.Cells[5 + i, 14].value = napomena20;

                        if (kolicina == 0)  // podbačaj radnika
                        {
                            worksheet4l.Rows.Cells[5 + i, 29].value = "";
                        }
                        else
                        {
                            worksheet4l.Rows.Cells[5 + i, 29].value = norm - kolicina;
                        }
                        if ((zadnjired==1) && !loop)
                        {
                            zadnjired = 0;
                            continue;
                        }

                        if (!loop)
                        {
                            zadnjired = 1;
                        }

                        i++;
                        hala2 = hala1;
                        napomena20 = napomena10;
                        smjena2 = smjena1;
                        proiz2 = proiz1;
                        radninalog2 = radninalog1;
                        linija2 = linija1;
                        tt2 = tt1;
                        cijena2 = cijena1;
                        norma = norma1;
                        norm = norm1;
                        erv = erv1;
                        normaukupp = normaukup1;
                        kolicina = kolicina1;
                        kupac2 = kupac1;
                        prvii = 0;
                        
                    }
                }
                
                cn.Close();
            }
            Console.WriteLine("Pregled po linijama  trenutno vrijeme " + DateTime.Now);

            smj = "6";  //   Štelanje
            dat13 = datLDP + " 06:00:00";
            dat23 = dat3 + " 05:59:00";

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
                workshstelanje.Rows.Cells[4, 3].value = workshstelanje.Rows.Cells[4, 3].value + "  " + datreps;

                while (reader.Read())
                {
                    workshstelanje.Rows.Cells[9 + i, 1].value = (reader["hala"].ToString());
                    workshstelanje.Rows.Cells[9 + i, 2].value = (reader["naziv"].ToString());
                    workshstelanje.Rows.Cells[9 + i, 3].value = (reader["aktivnost"].ToString());
                    workshstelanje.Rows.Cells[9 + i, 4].value = reader["pocetak"]; ;  // norma
                    workshstelanje.Rows.Cells[9 + i, 5].value = reader["kraj"]; ;  // komada                    
                    workshstelanje.Rows.Cells[9 + i, 6].value = reader["trajanje_minuta"];
                    workshstelanje.Rows.Cells[9 + i, 7].value = reader["izgubljeno"];
                    workshstelanje.Rows.Cells[9 + i, 8].value = reader["brojrn"];
                    workshstelanje.Rows.Cells[9 + i, 9].value = reader["napomena"];
                    workshstelanje.Rows.Cells[9 + i, 10].value = reader["username"];
                    workshstelanje.Rows.Cells[9 + i, 11].value = reader["vrstanarudzbe"];
                    workshstelanje.Rows.Cells[9 + i, 12].value = reader["nazivpar"];
                    i++;
                }

                cn.Close();
            }
            Console.WriteLine("Pregled izostanaka  trenutno vrijeme " + DateTime.Now);

            //workshstelanje.Rows.Cells[29, 3].value
            //            string fileName = @"C:\brisi\dsr20062017.xlsm";

            

            var fi = new FileInfo(fileName);
            if (fi.Exists) File.Delete(fileName);

            excel.Application.ActiveWorkbook.SaveAs(fileName);

            //file bez vrijednosti
            string fileNamebv  = @"c:\brisi\dsrv" + nuland + d2.Day.ToString() + nulanm + d2.Month.ToString() + d2.Year.ToString() + ".xlsx";
            string fileNamebv2 = @"c:\brisi\dsrv2_" + nuland + d2.Day.ToString() + nulanm + d2.Month.ToString() + d2.Year.ToString() + ".xlsx";

            fi = new FileInfo(fileNamebv);
            if (fi.Exists) File.Delete(fileNamebv);

            var fi2 = new FileInfo(fileNamebv2);          // horvat toni
            if (fi2.Exists) File.Delete(fileNamebv2);


            // ostavi samo samo sonu

            Worksheet worksheetbv2 = workbook.Worksheets.Item[1] as Worksheet;  // horvat toni

            worksheetbv2.Rows.Cells[30, 3].value = "";    // ukupna vrijednost
            // kaliona
            worksheetbv2.Rows.Cells[34, 11].value = "";    // realizacija
            worksheetbv2.Rows.Cells[35, 11].value = "";    // realizacija
            worksheetbv2.Rows.Cells[25, 17].value = "";    // vrijednost LGR

            Range startCell2 = worksheetbv2.Cells[43, 2];  // vrijednost repromaterijala
            Range endCell2 = worksheetbv2.Cells[43, 16];

            worksheetbv2.Range[startCell2, endCell2].Value = "";

            // ostavi samo sonu

            startCell2 = worksheetbv2.Cells[9, 6];  // vrijednost obrade u eur
            endCell2 = worksheetbv2.Cells[14, 7];
            worksheetbv2.Range[startCell2, endCell2].Value = "";

            startCell2 = worksheetbv2.Cells[16, 6];  // vrijednost obrade u eur
            endCell2 = worksheetbv2.Cells[22, 7];
            worksheetbv2.Range[startCell2, endCell2].Value = "";
            worksheetbv2.Range[startCell2, endCell2].Value = "";

            startCell2 = worksheetbv2.Cells[9,17 ];  // vrijednost sgr obrade u eur
            endCell2 = worksheetbv2.Cells[14, 17];
            worksheetbv2.Range[startCell2, endCell2].Value = "";

            startCell2 = worksheetbv2.Cells[16, 17];  // vrijednost sgr obrade u eur
            endCell2 = worksheetbv2.Cells[22, 17];
            worksheetbv2.Range[startCell2, endCell2].Value = "";

            Worksheet worksheetbvt2 = workbook.Worksheets.Item[2] as Worksheet;  // tjedni izvještaj Toni Horvat
            worksheetbvt2.Rows.Cells[30, 3].value = "";    // ukupna vrijednost u eur
            worksheetbvt2.Rows.Cells[34, 9].value = "";    // kaliona realizacija
            worksheetbvt2.Rows.Cells[25, 6].value = "";    // 
            worksheetbvt2.Rows.Cells[25, 7].value = "";    // 
            startCell2 = worksheetbvt2.Cells[9, 6];         // vrijednost obrade
            endCell2 = worksheetbvt2.Cells[14, 7];
            worksheetbvt2.Range[startCell2, endCell2].Value = "";
            startCell2 = worksheetbvt2.Cells[16, 6];         // vrijednost obrade
            endCell2 = worksheetbvt2.Cells[23, 7];
            worksheetbvt2.Range[startCell2, endCell2].Value = "";

            Worksheet worksheetbvm2 = workbook.Worksheets.Item[3] as Worksheet;  // mjesečni izvještaj Horvat Toni
            worksheetbvm2.Rows.Cells[30, 3].value = "";    // ukupna vrijednost u eur
            worksheetbvm2.Rows.Cells[34, 9].value = "";    // kaliona realizacija
            worksheetbvm2.Rows.Cells[25, 6].value = "";    // 
            worksheetbvm2.Rows.Cells[25, 7].value = "";    // 
            startCell2 = worksheetbvm2.Cells[9, 6];         // vrijednost obrade
            endCell2 = worksheetbvm2.Cells[14, 7];
            worksheetbvm2.Range[startCell2, endCell2].Value = "";
            startCell2 = worksheetbvm2.Cells[17, 6];         // vrijednost obrade
            endCell2 = worksheetbvm2.Cells[24, 7];
            worksheetbvm2.Range[startCell2, endCell2].Value = "";

            Worksheet worksheetbv32 = workbook.Worksheets.Item[4] as Worksheet;  // dnevni izvještaj p1 Horvat Toni
            worksheetbv32.Rows.Cells[30, 3].value = "";    // ukupna vrijednost u eur            
            worksheetbv32.Rows.Cells[25, 6].value = "";    // 
            worksheetbv32.Rows.Cells[25, 7].value = "";    // 
            startCell2 = worksheetbv32.Cells[9, 6];         // vrijednost obrade
            endCell2 = worksheetbv32.Cells[14, 7];
            worksheetbv32.Range[startCell2, endCell2].Value = "";
            startCell2 = worksheetbv32.Cells[17, 6];         // vrijednost obrade
            endCell2 = worksheetbv32.Cells[23, 7];
            worksheetbv32.Range[startCell2, endCell2].Value = "";

            Worksheet worksheetbv33 = workbook.Worksheets.Item[5] as Worksheet;  // dnevni izvještaj p3 Horvat Toni
            worksheetbv33.Rows.Cells[30, 3].value = "";    // ukupna vrijednost u eur            
            worksheetbv33.Rows.Cells[25, 6].value = "";    // 
            worksheetbv33.Rows.Cells[25, 7].value = "";    // 
            startCell2 = worksheetbv33.Cells[9, 6];         // vrijednost obrade
            endCell2 = worksheetbv33.Cells[14, 7];
            worksheetbv33.Range[startCell2, endCell2].Value = "";
            startCell2 = worksheetbv33.Cells[17, 6];         // vrijednost obrade
            endCell2 = worksheetbv33.Cells[23, 7];
            worksheetbv33.Range[startCell2, endCell2].Value = "";

            Worksheet worksheetbvlinija2 = workbook.Worksheets.Item[7] as Worksheet;  // po linijama, Horvat Toni
            startCell2 = worksheetbvlinija2.Cells[5, 9];  // kolona sa cijenama
            endCell2 = worksheetbvlinija2.Cells[310, 9];

            worksheetbvlinija2.Range[startCell2, endCell2].Value = "";

            startCell2 = worksheetbvlinija2.Cells[5, 10];  // kolona sa ostvarenim eurima
            endCell2 = worksheetbvlinija2.Cells[310, 10];

            worksheetbvlinija2.Range[startCell2, endCell2].Value = "";

            startCell2 = worksheetbvlinija2.Cells[5, 11];  // kolona sa planiranim eurima
            endCell2 = worksheetbvlinija2.Cells[310, 11];

            worksheetbvlinija2.Range[startCell2, endCell2].Value = "";

            startCell2 = worksheetbvlinija2.Cells[5, 13];  // kolona sa  razlikom u eurima
            endCell2 = worksheetbvlinija2.Cells[310, 13];

            worksheetbvlinija2.Range[startCell2, endCell2].Value = "";

            excel.Application.ActiveWorkbook.SaveAs(fileNamebv2);


            // voditelji pogonaetc.
            // ovaj puta brisi sve vrijednosti u EUR  - 
            Worksheet worksheetbv = workbook.Worksheets.Item[1] as Worksheet;

            worksheetbv.Rows.Cells[30, 3].value = "";    // ukupna vrijednost
            // kaliona
            worksheetbv.Rows.Cells[34, 11].value = "";    // realizacija
            worksheetbv.Rows.Cells[35, 11].value = "";    // realizacija
            worksheetbv.Rows.Cells[25, 17].value = "";    // vrijednost LGR

            Range startCell = worksheetbv.Cells[44, 2];  // vrijednost repromaterijala
            Range endCell = worksheetbv.Cells[44, 16];

            worksheetbv.Range[startCell, endCell].Value = "";

            startCell = worksheetbv.Cells[7, 6];         // vrijednost obrade
            endCell = worksheetbv.Cells[23, 7];
            worksheetbv.Range[startCell, endCell].Value = "";

            startCell = worksheetbv.Cells[7, 17];         // vrijednost sgr
            endCell = worksheetbv.Cells[23, 17];
            worksheetbv.Range[startCell, endCell].Value = "";

            Worksheet worksheetbvt = workbook.Worksheets.Item[2] as Worksheet;  // tjedni izvještaj
            worksheetbvt.Rows.Cells[30, 3].value = "";    // ukupna vrijednost u eur
            worksheetbvt.Rows.Cells[34, 9].value = "";    // kaliona realizacija
            worksheetbvt.Rows.Cells[25, 6].value = "";    // 
            worksheetbvt.Rows.Cells[25, 7].value = "";    // 
            startCell = worksheetbvt.Cells[7, 6];         // vrijednost obrade
            endCell   = worksheetbvt.Cells[23, 7];
            worksheetbvt.Range[startCell, endCell].Value = "";

            Worksheet worksheetbvm = workbook.Worksheets.Item[3] as Worksheet;  // mjesečni izvještaj
            worksheetbvm.Rows.Cells[30, 3].value = "";    // ukupna vrijednost u eur
            worksheetbvm.Rows.Cells[34, 9].value = "";    // kaliona realizacija
            worksheetbvm.Rows.Cells[25, 6].value = "";    // 
            worksheetbvm.Rows.Cells[25, 7].value = "";    // 
            startCell = worksheetbvm.Cells[7, 6];         // vrijednost obrade
            endCell   = worksheetbvm.Cells[23, 7];
            worksheetbvm.Range[startCell, endCell].Value = "";

            Worksheet worksheetbv31 = workbook.Worksheets.Item[4] as Worksheet;  // dnevni izvještaj p1
            worksheetbv31.Rows.Cells[30, 3].value = "";    // ukupna vrijednost u eur            
            worksheetbv31.Rows.Cells[25, 6].value = "";    // 
            worksheetbv31.Rows.Cells[25, 7].value = "";    // 
            startCell = worksheetbv31.Cells[7, 6];         // vrijednost obrade
            endCell = worksheetbv31.Cells[23, 7];
            worksheetbv31.Range[startCell, endCell].Value = "";
            
            
            Worksheet worksheetbv3 = workbook.Worksheets.Item[5] as Worksheet;  // dnevni izvještaj p3
            worksheetbv3.Rows.Cells[30, 3].value = "";    // ukupna vrijednost u eur            
            worksheetbv3.Rows.Cells[25, 6].value = "";    // 
            worksheetbv3.Rows.Cells[25, 7].value = "";    // 
            startCell = worksheetbv3.Cells[7, 6];         // vrijednost obrade
            endCell   = worksheetbv3.Cells[23, 7];
            worksheetbv3.Range[startCell, endCell].Value = "";

            
            Worksheet worksheetbvlinija = workbook.Worksheets.Item[7] as Worksheet;  // po linijama
            startCell = worksheetbvlinija.Cells[5, 9];  // kolona sa cijenama
            endCell = worksheetbvlinija.Cells[290, 9];

            worksheetbvlinija.Range[startCell, endCell].Value = "";

            startCell = worksheetbvlinija.Cells[5, 10];  // kolona sa ostvarenim eurima
            endCell = worksheetbvlinija.Cells[290, 10];

            worksheetbvlinija.Range[startCell, endCell].Value = "";

            startCell = worksheetbvlinija.Cells[5, 11];  // kolona sa planiranim eurima
            endCell = worksheetbvlinija.Cells[290, 11];

            worksheetbvlinija.Range[startCell, endCell].Value = "";

            startCell = worksheetbvlinija.Cells[5, 13];  // kolona sa  razlikom u eurima
            endCell = worksheetbvlinija.Cells[290, 13];

            worksheetbvlinija.Range[startCell, endCell].Value = "";

            excel.Application.ActiveWorkbook.SaveAs(fileNamebv);

            ///////////////////////////////

            Console.WriteLine("Snimljen file ddmmyyyy   trenutno vrijeme " + DateTime.Now);
            //Console.ReadKey();
            //workbook.Close(false);
            //excel.Application.Quit();
            //excel.Quit()
            //workbook.Close(true, Type.Missing, Type.Missing);
            workbook.Close(false, Type.Missing, Type.Missing);

            excel.Application.Quit();
            excel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            workshstelanje = null;
            workbook = null;
            app = null;
            int test = 10;

            MailMessage mail = new MailMessage("gasparic.s@feroimpex.hr", "gasparic.s@feroimpex.hr,srecckog@gmail.com");

    if (test==0)
            mail = new MailMessage("gasparic.s@feroimpex.hr", "legac.b@feroimpex.hr,legac.h@feroimpex.hr,legac.z@feroimpex.hr,horvatic.v@feroimpex.hr,info@feroimpex.hr,gasparic.s@feroimpex.hr");            
    
            SmtpClient client = new SmtpClient();
            client.Port = 25;
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            client.UseDefaultCredentials = false;
            client.Credentials = new System.Net.NetworkCredential("gasparic.s@feroimpex.hr", "gasparic1");
            client.Host = "mail.feroimpex.hr";
            mail.Subject = "DPR " + datreps;
            mail.Body = "Daily production report for " + datreps;

            Attachment attachment = new Attachment(fileName, System.Net.Mime.MediaTypeNames.Application.Octet);
            System.Net.Mime.ContentDisposition disposition = attachment.ContentDisposition;
            disposition.CreationDate = File.GetCreationTime(fileName);
            disposition.ModificationDate = File.GetLastWriteTime(fileName);
            disposition.ReadDate = File.GetLastAccessTime(fileName);
            disposition.FileName = Path.GetFileName(fileName);
            disposition.Size     = new FileInfo(fileName).Length;
            disposition.DispositionType = System.Net.Mime.DispositionTypeNames.Attachment;

            mail.Attachments.Add(attachment);
            client.Send(mail);
            client.Dispose();

if (test == 0)
            mail = new MailMessage("gasparic.s@feroimpex.hr", "jancic.d@feroimpex.hr,cakanic.s@feroimpex.hr,kicin.d@feroimpex.hr,grgecic.d@feroimpex.hr,gasparic.s@feroimpex.hr,srecckog@gmail.com,stefanac.d@feroimpex.hr,darapi.i@feroimpex.hr,janjcic.d@feroimpex.hr, hren.h@feroimpex.hr, igrec.m@feroimpex.hr,biscan.s@feroimpex.hr,gradiski.b@feroimpex.hr,vladic.p@feroimpex.hr,korpak.d@feroimpex.hr,barbic.i@feroimpex.hr,golesic.m@feroimpex.hr");
            // ovaj puta šalji bez vrijednosti u EUR

            //mail = new MailMessage("gasparic.s@feroimpex.hr", "gasparic.s@feroimpex.hr,srecckog@gmail.com");

            client = new SmtpClient();
            client.Port = 25;
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            client.UseDefaultCredentials = false;
            client.Credentials = new System.Net.NetworkCredential("gasparic.s@feroimpex.hr", "gasparic1");
            client.Host = "mail.feroimpex.hr";

            mail.Subject = "_DPR " + datreps;
            mail.Body = "_Daily production report for " + datreps;

            Attachment attachment2 = new Attachment(fileNamebv, System.Net.Mime.MediaTypeNames.Application.Octet);
            System.Net.Mime.ContentDisposition disposition2 = attachment2.ContentDisposition;
            disposition2.CreationDate = File.GetCreationTime(fileNamebv);
            disposition2.ModificationDate = File.GetLastWriteTime(fileNamebv);
            disposition2.ReadDate = File.GetLastAccessTime(fileNamebv);
            disposition2.FileName = Path.GetFileName(fileNamebv);
            disposition2.Size = new FileInfo(fileNamebv).Length;
            disposition2.DispositionType = System.Net.Mime.DispositionTypeNames.Attachment;

            mail.Attachments.Remove(attachment);
            mail.Attachments.Add(attachment2);
            client.Send(mail);
            client.Dispose();

            if (IndexCNC == 1)
            {
                mail = new MailMessage("gasparic.s@feroimpex.hr", "srecckog@gmail.com,gasparic.s@feroimpex.hr");
                //mail = new MailMessage("gasparic.s@feroimpex.hr", "gasparic.s@feroimpex.hr,srecckog@gmail.com");

                client = new SmtpClient();
                client.Port = 25;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.UseDefaultCredentials = false;
                client.Credentials = new System.Net.NetworkCredential("gasparic.s@feroimpex.hr", "gasparic1");
                client.Host = "mail.feroimpex.hr";

                mail.Subject = " !!! INDEX LINIJA IMA OBRADU TOKARENJA , promjeni store " + datreps;
                mail.Body = " !!! INDEX LINIJA IMA OBRADU TOKARENJA , promjeni store " + datreps;

                //Attachment attachment31 = new Attachment(fileNamebv2, System.Net.Mime.MediaTypeNames.Application.Octet);
                //System.Net.Mime.ContentDisposition disposition31 = attachment3.ContentDisposition;
                //disposition3.CreationDate = File.GetCreationTime(fileNamebv2);
                //disposition3.ModificationDate = File.GetLastWriteTime(fileNamebv2);
                //disposition3.ReadDate = File.GetLastAccessTime(fileNamebv2);
                //disposition3.FileName = Path.GetFileName(fileNamebv2);
                //disposition3.Size = new FileInfo(fileNamebv2).Length;
                //disposition3.DispositionType = System.Net.Mime.DispositionTypeNames.Attachment;

                mail.Attachments.Remove(attachment2);
                //mail.Attachments.Add(attachment3);
                client.Send(mail);
                client.Dispose();
            }


            if (test==0)
            mail = new MailMessage("gasparic.s@feroimpex.hr", "horvat.t@feroimpex.hr,gasparic.s@feroimpex.hr");
            //mail = new MailMessage("gasparic.s@feroimpex.hr", "gasparic.s@feroimpex.hr,srecckog@gmail.com");

            client = new SmtpClient();
            client.Port = 25;
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            client.UseDefaultCredentials = false;
            client.Credentials = new System.Net.NetworkCredential("gasparic.s@feroimpex.hr", "gasparic1");
            client.Host = "mail.feroimpex.hr";

            mail.Subject = "_DPR2 " + datreps;
            mail.Body = "_Daily production report for " + datreps;

            Attachment attachment3 = new Attachment(fileNamebv2, System.Net.Mime.MediaTypeNames.Application.Octet);
            System.Net.Mime.ContentDisposition disposition3 = attachment3.ContentDisposition;
            disposition3.CreationDate = File.GetCreationTime(fileNamebv2);
            disposition3.ModificationDate = File.GetLastWriteTime(fileNamebv2);
            disposition3.ReadDate = File.GetLastAccessTime(fileNamebv2);
            disposition3.FileName = Path.GetFileName(fileNamebv2);
            disposition3.Size = new FileInfo(fileNamebv2).Length;
            disposition3.DispositionType = System.Net.Mime.DispositionTypeNames.Attachment;

            mail.Attachments.Remove(attachment2);
            mail.Attachments.Add(attachment3);
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
