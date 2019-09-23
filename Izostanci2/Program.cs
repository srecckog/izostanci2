using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.IO;
using System.Net.Mail;
using System.Reflection;
using System.Diagnostics;

namespace Izostanci2
{
    class Program
    {
        static void Main(string[] args)
        {

            //Console.WriteLine("11111111111111111111111" + DateTime.Now);
            //Console.ReadKey();

            string connectionString = @"Data Source=192.168.0.5;Initial Catalog=FeroApp;User ID=sa;Password=AdminFX9.";
            string connectionStringRFIND = @"Data Source=192.168.0.3;Initial Catalog=RFIND;User ID=sa;Password=AdminFX9.";

            string c7 = "", c8 = "", c9 = "", c10 = "", c11 = "", c12 = "", c13 = "", c14 = "", c15 = "0", c16 = "", c17 = "", c19 = "", c21 = ""; // 
            string c7do = "", c8do = "", c9do = "", c10do = "", c11do = "", c12do = "", c13do = "", c14do = "", c15do = "", c16do = "", c17do = "", c19do = "", c21do = ""; // 
            double i7 = 0.0, i8 = 0.0, i9 = 0.0, i10 = 0.0, i11 = 0.0, i12 = 0.0, i13 = 0.0, i14 = 0.0, i15 = 0.0, i16 = 0.0, i17 = 0.0, i18 = 0.0, i19 = 0.0, i21 = 0.0, i22 = 0; // 
            double f7 = 0.0, f8 = 0.0, f9 = 0.0, f10 = 0.0, f11 = 0.0, f12 = 0.0, f13 = 0.0, f14 = 0.0, f15 = 0.0, f16 = 0.0, f17 = 0.0, f18 = 0.0, f19 = 0.0, f21 = 0.0, f22 = 0; // 
            double g7 = 0.0, g8 = 0.0, g9 = 0.0, g10 = 0.0, g11 = 0.0, g12 = 0.0, g13 = 0.0, g14 = 0.0, g15 = 0.0, g16 = 0.0, g17 = 0.0, g18 = 0.0, g19 = 0.0, g21 = 0.0, g22 = 0; // 
            string j7 = "0", j8 = "0", j9 = "0", j10 = "0", j11 = "0", j12 = "0", j13 = "0", j14 = "0", j15 = "0", j16 = "0", j17 = "0", j19 = "0", j21 = "0"; // 
            string sql1 = "";
            string dat1 = "2017-06-20", dat2 = "2017-06-20", dat3 = "", dat3p = "";


            string baza1 = "fxsap.dbo.plansatirada";
            //baza1 = "rfind.dbo.plansatirada2";
            int test = 0;   // 10-test, 0 - live

            // Process.Start("C:\\brisi\\_raspored djelatnika21.xlsm");

            DateTime d1 = new DateTime(2017, 6, 20);
            DateTime d10 = new DateTime(2017, 7, 27);
            //DateTime d2 = DateTime.Now.AddDays( -1 )  ; // smjena 2
            //d2 = d2.AddHours(-8);

            DateTime d3 = DateTime.Now;
            DateTime d3p = DateTime.Now;
            DateTime d2 = DateTime.Now;
            if (d2.Hour < 14 && 1 == 1)
            {
                //d2 = d2.AddDays(-1);
                //         d2 = d2.AddHours(-7);   // brisi
                d3p = d2;
            }

            //     druga smjena od prethodnog dana
            //              d2 = d2.AddHours(15);


            // test on date

            //d2 = new DateTime(2018, 8, 5);
            //d2 = d2.AddHours(0);

            //


            d3 = d2;

            string dan1 = d2.Day.ToString();
            string m1 = d2.Month.ToString();
            string g1 = d2.Year.ToString();
            int smjenaz = 2;                        // smjena ???

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

            int brojsati = 7;
            if ((d2.DayOfWeek == DayOfWeek.Sunday) || (d2.DayOfWeek == DayOfWeek.Saturday))
            {
                brojsati = 5;
            }

            DateTime sunday = input.AddDays(delta);
            //monday= new DateTime(2017, 7, 31);
            sunday = monday.AddDays(6);
            //sunday = d2;

            string mm1 = "";
            if (d2.Month <= 9)
                mm1 = "0";

            dat1 = d2.Year.ToString() + '-' + mm1 + d2.Month.ToString() + '-' + d2.Day.ToString();
            DateTime d30 = DateTime.Now;
            mm1 = "";
            if (d30.Month <= 9)
                mm1 = "0";
            string dats = d30.Year.ToString() + '-' + mm1 + d30.Month.ToString() + '-' + d30.Day.ToString();  // današnji datum
            dat2 = dat1;
            mm1 = "";
            if (d3.Month <= 9)
                mm1 = "0";
            dat3 = d3.Year.ToString() + '-' + mm1 + d3.Month.ToString() + '-' + d3.Day.ToString();
            dat3p = d3p.Year.ToString() + '-' + d3p.Month.ToString() + '-' + d3p.Day.ToString();
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
            string dat30 = dnuland + d3.Day.ToString() + '.' + dnulanm + d3.Month.ToString() + '.' + d3.Year.ToString();   // današnji datum
            string dat13 = dat1 + " 6:00:00";
            DateTime d23 = d2.AddDays(1);
            //string dat23 = d2.Year.ToString() + '-' + d2.Month.ToString() + '-' + d2.Day.ToString() + " 6:00:00";

            mm1 = "";
            if (d2.Month <= 9)
                mm1 = "0";

            string dat23 = d2.Year.ToString() + '-' + mm1 + d2.Month.ToString() + '-' + d2.Day.ToString();

            //string fileName = @"L:\izvještaji\dsr\dprs" + nuland + d2.Day.ToString() + nulanm + d2.Month.ToString() + d2.Year.ToString() + ".xlsm";

            string smjenanaziv = "", smj = "";
            DateTime datrep = d3; ;

            //            sql1 = "rfind.dbo.izostanci '" + dat1 + "',1";

            string fileNameIzo = @"l:\izvještaji\dsr\izostanci_" + nuland + d2.Day.ToString() + nulanm + d2.Month.ToString() + d2.Year.ToString() + "_" + smj + ".xlsm";
            string fileNameNepl = @"l:\izvještaji\dsr\izostanci2_" + nuland + d2.Day.ToString() + nulanm + d2.Month.ToString() + d2.Year.ToString() + "_" + smj + ".xlsm";

            if (datrep.Day <= 9)
            {
                dnuland = "0";
            }
            if (datrep.Month <= 9)
            {
                dnulanm = "0";
            }
            string dat10sp = dat10;
            string dat1sp = dat1;
            string dat2sp = dat2;
            string datreps = dnuland + datrep.Day.ToString() + '.' + dnulanm + datrep.Month.ToString() + '.' + datrep.Year.ToString();   // današnji datum
            string datLDP = dat1;
            Application excel = new Application();
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();

            //Workbook workbook = excel.Workbooks.Open(@"c:\izvještaji\dsr\dprs_template1.xlsm", ReadOnly: false, Editable: true);
            // Workbook workbook = excel.Workbooks.Open(@"c:\brisi\dprs_template1.xlsm", ReadOnly: false, Editable: true);
            //            Workbook workbook = excel.Workbooks.Open(@"c:\brisi\dprs_template1601.xlsm", ReadOnly: false, Editable: true);

            //            Workbook workbook = excel.Workbooks.Open(@"c:\brisi\dprs_template1304.xlsm", ReadOnly: false, Editable: true);
            // template
            Workbook workbook = excel.Workbooks.Open(@"c:\brisi\izostanci_1607.xlsm", ReadOnly: false, Editable: true);

            dat10 = dat10sp;
            dat1 = dat1sp;
            dat2 = dat2sp;

            Worksheet worksheetIZO = workbook.Worksheets.Item[1] as Worksheet;

            Worksheet worksheetPlansatirada = workbook.Worksheets.Item[5] as Worksheet;

            int i = 1, j = 1, k = 1, ir = 1;

            DateTime jucer1 = DateTime.Now.AddDays(-1);
            //////////////////////////////////////////////////////////////
           jucer1 = new DateTime(2019,  9, 22 , 12, 12,0);
            //////////////////////////////////////////////////////////////

            int dayofweek1 = (int)jucer1.DayOfWeek;

            string mjes1 = jucer1.Month.ToString();
            string god1 = jucer1.Year.ToString();

            string danj = jucer1.Day.ToString();
            string dj, mj;
            dj = danj.Trim();
            if (danj.Trim().Length == 1)
            {
                dj = "0" + danj.Trim();
            }

            mj = jucer1.Month.ToString();
            string mj0 = mj;

            if (jucer1.Month <= 9)
            {
                mj = "0" + jucer1.Month.ToString();
            }
            string djucer = jucer1.Year.ToString() + "-" + mj + "-" + danj;
            Console.WriteLine("Od datuma: " + djucer + " - " + djucer + " trenutno vrijeme " + DateTime.Now);

            dat1 = djucer;

            if (1 == 1)
            {
                // update plansatirada
                // 1.korak, mt=700,716 and unešeno od štelera

                //            sql1 = "select * from(select x1.*, p.dosao, p.otisao, p.Ukupno_minuta, p.RadnoMjesto, case when dosao is null then 'Nije zaplaniran' when year(dosao) = 1900 then 'Nije se prijavio' when year(dosao) = 1900 then 'Nije se prijavio' end napomena, case when x1.smjena = 3 and x1.satiradaradnika >= 8 then '1n' when x1.smjena = 2 and x1.satiradaradnika >= 8 then '1p' when x1.smjena = 1 and x1.satiradaradnika >= 8 then '1j' when x1.smjena = 3 and x1.satiradaradnika < 8 and x1.satiradaradnika > 0 then('-' + cast((7.0 - x1.satiradaradnika) as varchar(12)) + 'n') when x1.smjena = 2 and x1.satiradaradnika < 8 and x1.satiradaradnika > 0  then('-' + cast((7.0 - x1.satiradaradnika) as varchar(12)) + 'p') when x1.satiradaradnika = 0 then '0e' when x1.smjena = 1 and x1.satiradaradnika < 8 and x1.satiradaradnika > 0 then('-' + cast((7.0 - x1.satiradaradnika) as varchar(12)) + 'j') end oznaka from( select e.radnik, e.id_radnika, e.vrsta, e.firma, e.datum, e.hala, e.smjena, sum(e.satiradaradnika) satiradaradnika, id FROM feroapp.dbo.EvidNormiRada('" + djucer+ "', '" + djucer +"') e " +
                //                   "inner join rfind.dbo.radnici_ r on r.id_radnika = e.id_radnika group by e.radnik, e.id_radnika, e.vrsta, e.firma, e.datum, e.hala, e.smjena, id ) x1 left join rfind.dbo.pregledvremena p on p.IDRadnika = x1.id and p.datum = x1.datum and x1.smjena = p.smjena ) x2  order by x2.radnik,x2.smjena";

                //sql1 = "select * from(select x1.*, p.dosao, p.otisao, p.Ukupno_minuta, p.RadnoMjesto, case when dosao is null then 'Nije zaplaniran' when year(dosao) = 1900 then 'Nije se prijavio' when year(dosao) = 1900 then 'Nije se prijavio' end napomena, case when x1.smjena = 3 and x1.satiradaradnika >= 8 then '1n' when x1.smjena = 2 and x1.satiradaradnika >= 8 then '1p' when x1.smjena = 1 and x1.satiradaradnika >= 8 then '1j' when x1.smjena = 3 and x1.satiradaradnika < 8 and x1.satiradaradnika > 0 then('-' + cast((7.0 - x1.satiradaradnika) as varchar(12)) + 'n') when x1.smjena = 2 and x1.satiradaradnika < 8 and x1.satiradaradnika > 0  then('-' + cast((7.0 - x1.satiradaradnika) as varchar(12)) + 'p') when x1.satiradaradnika = 0 then '0e' when x1.smjena = 1 and x1.satiradaradnika < 8 and x1.satiradaradnika > 0 then('-' + cast((7.0 - x1.satiradaradnika) as varchar(12)) + 'j') end oznaka from( select e.radnik, e.id_radnika, e.vrsta, e.firma, e.datum, e.hala, e.smjena, sum(e.satiradaradnika) satiradaradnika, id, PomocniRadnik FROM feroapp.dbo.EvidNormiRada('" + djucer + "', '" + djucer + "') e " +
                //       "inner join rfind.dbo.radnici_ r on r.id_radnika = e.id_radnika group by e.radnik, e.id_radnika, e.vrsta, e.firma, e.datum, e.hala, e.smjena, id, PomocniRadnik union all select e.pomocniradnik radnik, e.id_radnika2, e.vrsta, e.firma, e.datum, e.hala, e.smjena, sum(e.SatiRadaPomocnogRadnika) satiradaradnika, id, PomocniRadnik FROM feroapp.dbo.EvidNormiRada('" + djucer + "', '" + djucer + "') e inner join rfind.dbo.radnici_ r on r.id_radnika = e.id_radnika2 where pomocniradnik is not null and  pomocniradnik != '' group by e.radnik, e.id_radnika2, e.vrsta, e.firma, e.datum, e.hala, e.smjena, id, PomocniRadnik ) x1   left join rfind.dbo.pregledvremena p on p.IDRadnika = x1.id and p.datum = x1.datum and x1.smjena = p.smjena  ) x2 order by x2.radnik,x2.smjena ";

                //                sql1 = "select r.prezime,r.ime,x2.* from(select x1.*, p.dosao, p.otisao, p.Ukupno_minuta, p.RadnoMjesto, case when dosao is null then 'Nije zaplaniran' when year(dosao) = 1900 then 'Nije se prijavio' when year(dosao) = 1900 then 'Nije se prijavio' end napomena, case when x1.smjena = 3 and x1.satiradaradnika >= 8 then '1n' when x1.smjena = 2 and x1.satiradaradnika >= 8 then '1p' when x1.smjena = 1 and x1.satiradaradnika >= 8 then '1j' when x1.smjena = 3 and x1.satiradaradnika < 8 and x1.satiradaradnika > 0 then('-' + cast((7.0 - x1.satiradaradnika) as varchar(12)) + 'n') when x1.smjena = 2 and x1.satiradaradnika < 8 and x1.satiradaradnika > 0  then('-' + cast((7.0 - x1.satiradaradnika) as varchar(12)) + 'p') when x1.satiradaradnika = 0 then '0e' when x1.smjena = 1 and x1.satiradaradnika < 8 and x1.satiradaradnika > 0 then('-' + cast((7.0 - x1.satiradaradnika) as varchar(12)) + 'j') end oznaka from(select x11.* from(select x11.*, id from(select e.id_radnika, e.firma, e.datum, e.hala, e.smjena, sum(e.satiradaradnika) satiradaradnika FROM feroapp.dbo.EvidNormiRada('" + djucer + "', '" + djucer + "') e " +
                //                       "group by e.id_radnika, e.firma, e.datum, e.hala, e.smjena ) x11  inner join rfind.dbo.radnici_ r on r.id_radnika = x11.id_radnika  ) x11  union all select x12.* from( select x12.*, id from( select e.id_radnika2 id_radnika, e.firma, e.datum, e.hala, e.smjena, sum(e.satiradapomocnogradnika) satiradaradnika  FROM feroapp.dbo.EvidNormiRada('" + djucer + "', '" + djucer + "') e  where pomocniradnik is not null and  pomocniradnik != ''  group by e.id_radnika2, e.firma, e.datum, e.hala, e.smjena  ) x12  inner join rfind.dbo.radnici_ r on r.id_radnika = x12.id_radnika  ) x12 ) x1 left join rfind.dbo.pregledvremena p on p.IDRadnika = x1.id and p.datum = x1.datum and x1.smjena = p.smjena  ) x2 left join rfind.dbo.radnici_ r on r.id = x2.id order by x2.id_radnika,x2.smjena";

                //sql1 = "select r.prezime,r.ime,x2.* from(select x1.*, p.dosao, p.otisao, p.Ukupno_minuta, p.RadnoMjesto,case when dosao is null then 'Nije zaplaniran' when year(dosao) = 1900 then 'Nije se prijavio' when year(dosao) = 1900 then 'Nije se prijavio' end napomena, " +
                //       "case when DATEPART(WEEKDAY, '"+djucer+"') != 6 and  x1.smjena = 3 and x1.satiradaradnika >= 8 then '1n' when DATEPART(WEEKDAY, '"+djucer+"') not in (6) and x1.smjena = 2 and x1.satiradaradnika >= 8 then '1p' when DATEPART(WEEKDAY, '"+djucer+"') not in (6) and x1.smjena = 1 and x1.satiradaradnika >= 8 then '1j' when DATEPART(WEEKDAY, '"+djucer+"') not in (6) and x1.smjena = 3 and x1.satiradaradnika < 8 and x1.satiradaradnika > 0 then('-' + cast((7.0 - x1.satiradaradnika) as varchar(12)) + 'n') when DATEPART(WEEKDAY, '"+djucer+"') not in (6) and x1.smjena = 2 and x1.satiradaradnika < 8 and x1.satiradaradnika > 0 then('-' + cast((7.0 - x1.satiradaradnika) as varchar(12)) + 'p') when DATEPART(WEEKDAY, '"+djucer+"') not in (6) and x1.smjena = 1 and x1.satiradaradnika < 8 and x1.satiradaradnika > 0 then('-' + cast((7.0 - x1.satiradaradnika) as varchar(12)) + 'j') when DATEPART(WEEKDAY, '"+djucer+"') = 6 and  x1.smjena = 2 and x1.satiradaradnika >= 8 then '3p' " +
                //       " when DATEPART(WEEKDAY, '"+djucer+"') = 6 and  x1.smjena = 1 and x1.satiradaradnika >= 8 then '3j' when DATEPART(WEEKDAY, '"+djucer+"') = 6 and  x1.smjena = 2 and x1.satiradaradnika < 8 and x1.satiradaradnika > 0 then('-' + cast((5.0 - x1.satiradaradnika) as varchar(12)) + 'p') when DATEPART(WEEKDAY, '"+djucer+ "') = 6 and  x1.smjena = 1 and x1.satiradaradnika < 8 and x1.satiradaradnika > 0 then('-' + cast((5.0 - x1.satiradaradnika) as varchar(12)) + 'j') when x1.satiradaradnika = 0 and ukupno_minuta=0 then '0e' when x1.satiradaradnika = 0 and ukupno_minuta>59 then '0X' end oznaka from( select x11.* from(  select x11.*, id from( select e.id_radnika, e.firma, e.datum, e.hala, e.smjena, sum(e.satiradaradnika) satiradaradnika FROM feroapp.dbo.EvidNormiRada('" + djucer+"', '"+djucer+"') e group by e.id_radnika, e.firma, e.datum, e.hala, e.smjena ) x11  inner join rfind.dbo.radnici_ r on r.id_radnika = x11.id_radnika ) x11 union all  select x12.* from( select x12.*, id from(          select e.id_radnika2 id_radnika, e.firma, e.datum, e.hala, e.smjena, sum(e.SatiRadaPomocnogRadnika) satiradaradnika  FROM feroapp.dbo.EvidNormiRada('"+djucer+"', '"+djucer+"') e where pomocniradnik is not null and  pomocniradnik != '' group by e.id_radnika2, e.firma, e.datum, e.hala, e.smjena ) x12 inner join rfind.dbo.radnici_ r on r.id_radnika = x12.id_radnika " +
                //       " ) x12 ) x1 left join rfind.dbo.pregledvremena p on p.IDRadnika = x1.id and p.datum = x1.datum and x1.smjena = p.smjena  ) x2 inner join rfind.dbo.radnici_ r on r.id_radnika = x2.id_radnika  order by x2.id_radnika,x2.smjena";

                string danxx = "Dan" + dj;
                int praznik = 0;
                worksheetPlansatirada.Rows.Cells[i, 6].value = danxx;

                // 0. korak praznici
                using (SqlConnection cn = new SqlConnection(connectionStringRFIND))  // praznici
                {
                    sql1 = "update " + baza1 + " set " + danxx + "='0y' where '" + djucer + "' in ( select datum from rfind.dbo.praznici) and " + danxx + " is null and mjesec=" + mj0 + " and godina=" + god1;    //   feroapp.dbo.evidnormirada( dat1 , dat2 )
                    cn.Open();
                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);
                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    if (reader.HasRows)
                    {
                        praznik = 1;
                    }
                    cn.Close();
                }

                //1.korak evidenormirada, upis od strane štelera

                using (SqlConnection cn = new SqlConnection(connectionStringRFIND))
                {

                    sql1 = "rfind.dbo.satniceoznake1 '" + djucer + "'";    //   feroapp.dbo.evidnormirada( dat1 , dat2 )
                    cn.Open();
                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string id1, firma1, sql2, ozn0, ozn1, oznold = "", firma10 = "", id10 = "", smjena1 = "";
                    int prviputa = 1, jj = 0, ukminuta1 = 0;


                    while (reader.Read())
                    {
                        if (reader["datum"] != DBNull.Value)
                        {

                            if ((reader["id"].ToString() == id10) && (reader["firma"].ToString() == firma10))
                            {
                                prviputa = 0;
                            }
                            else
                            {
                                prviputa = 1;
                                jj = 0;
                                oznold = "";
                            }

                            id1 = reader["id"].ToString();
                            if (id1 == "1485")
                            {
                                int hh = 0;
                            }
                            int intid1 = (int.Parse)(id1);

                            if (intid1 > 8000 && intid1 < 9000)
                            {
                                intid1 = intid1 - 8000;
                                id1 = intid1.ToString();
                            }
                            if (intid1 == 1269)
                            {
                                int uu = 0;
                            }
                            firma1 = reader["firma"].ToString();
                            smjena1 = reader["smjena"].ToString();
                            if (reader["ukupno_minuta"] != DBNull.Value)
                            {
                                ukminuta1 = (int.Parse)(reader["ukupno_minuta"].ToString());
                            }
                            else
                            {
                                ukminuta1 = -99;
                            }
                            // subota 3 smjena i nedjelja se ne upisuju u plansatirada
                            if ((smjena1 == "3" && dayofweek1 == 6) || (dayofweek1 == 0))
                            {
                                continue;
                            }

                            //m1 = reader["mjesec"].ToString();
                            //g1 = reader["godina"].ToString();
                            ozn1 = reader["oznaka"].ToString().Replace(".00", "");
                            ozn1 = ozn1.Replace("-0", "0");
                            //ozn1 = ozn1.Replace("0X", "");
                            id10 = id1;
                            firma10 = firma1;

                            //if (firma1 == "FX")
                            //{
                            //    firma1 = "3";
                            //}
                            //else
                            //{
                            //    firma1 = "1";
                            //}

                            sql2 = " select * from " + baza1 + " where mjesec=" + mjes1 + " and godina=" + god1 + " and radnikid= " + id1 + " and firma=" + firma1;

                            using (SqlConnection cn2 = new SqlConnection(connectionStringRFIND))
                            {

                                cn2.Open();
                                SqlCommand sqlCommand2 = new SqlCommand(sql2, cn2);
                                SqlDataReader reader2 = sqlCommand2.ExecuteReader();
                                string id2, firma2;

                                while (reader2.Read())
                                {

                                    if ((reader2[danxx] == DBNull.Value) || (reader2[danxx].ToString().TrimEnd() == ""))
                                    {// ako je null

                                        using (SqlConnection cn20 = new SqlConnection(connectionStringRFIND))
                                        {
                                            cn20.Open();
                                            string sql20 = "update " + baza1 + " set " + danxx + "= '" + ozn1 + "' where  mjesec=" + mjes1 + " and godina=" + god1 + " and radnikid= " + id1 + " and firma='" + firma1 + "'";
                                            SqlCommand sqlc20 = new SqlCommand(sql20, cn20);
                                            SqlDataReader reader20 = sqlc20.ExecuteReader();
                                            cn20.Close();
                                        }
                                        jj++;
                                        oznold = ozn1;


                                    }
                                    else
                                    { // ako je već upisano
                                        string postojecaoznaka = reader2[danxx].ToString();
                                        string radnikk = reader2["ime"].ToString();

                                        if (prviputa == 1)
                                        {// ne diraj ako je netko već nešto upisao prije 
                                            int zz = 0;
                                            oznold = ozn1;

                                        }
                                        else
                                        {
                                            if ((oznold != ozn1) && jj > 0) // ako je na početku bilo prazno
                                            {
                                                if (id1 == "159")
                                                {
                                                    int hh = 0;
                                                }
                                                oznold = ozn1;
                                                ozn1 = reader2[danxx].ToString() + ":" + ozn1;

                                                int countj = 0;
                                                if (ozn1.Split('j').Length > 0)
                                                {
                                                    countj = ozn1.Split('j').Length - 1;
                                                }
                                                int countp = 0;
                                                if (ozn1.Split('p').Length > 0)
                                                {
                                                    countp = ozn1.Split('p').Length - 1;
                                                }
                                                int countn = 0;
                                                if (ozn1.Split('n').Length > 0)
                                                {
                                                    countn = ozn1.Split('n').Length - 1;
                                                }
                                                int check1 = 0;
                                                string ozn11 = "";
                                                if (ukminuta1 > 465 && ukminuta1 < 525)
                                                {

                                                    if (countj == 2)
                                                        ozn11 = "1j";
                                                    if (countp == 2)
                                                        ozn11 = "1p";
                                                    if (countn == 2)
                                                        ozn11 = "1n";

                                                    check1 = 1;
                                                }
                                                string brojsatii = (ukminuta1 / 60).ToString();
                                                if ((ukminuta1 > 540) && !provjerasati(ozn1, brojsatii,brojsati).Contains("Ok")  )       // radio više od 8 sati ?
                                                {
                                                    i++;
                                                    worksheetPlansatirada.Rows.Cells[i, 1].value = firma1;
                                                    worksheetPlansatirada.Rows.Cells[i, 2].value = id1;
                                                    worksheetPlansatirada.Rows.Cells[i, 3].value = reader2["ime"].ToString();
                                                    worksheetPlansatirada.Rows.Cells[i, 4].value = reader2["sifrarm"].ToString();
                                                    worksheetPlansatirada.Rows.Cells[i, 5].value = reader2["mt"].ToString();
                                                    worksheetPlansatirada.Rows.Cells[i, 6].value = ozn1;
                                                    worksheetPlansatirada.Rows.Cells[i, 7].value = (ukminuta1 / 60).ToString();
                                                    //worksheetPlansatirada.Rows.Cells[i, 8].value = reader2["RadnoMjesto"].ToString();
                                                    worksheetPlansatirada.Rows.Cells[i, 11].value = "13. radio više od 8 sati >> ";
                                                }

                                                if (countj > 1 || countp > 1 || countn > 1)
                                                {
                                                    Console.WriteLine("1.Potrebna provjera za dan " + danxx + " Ime=" + radnikk + "   id1 = " + id1 + " Firma = " + firma1 + " ozn1 " + ozn1 + " odradio ukupno minuta " + ukminuta1.ToString() + " >> check1 " + check1.ToString() + " ozn11 " + ozn11);
                                                    //Console.ReadKey();
                                                    i++;
                                                    worksheetPlansatirada.Rows.Cells[i, 1].value = firma1;
                                                    worksheetPlansatirada.Rows.Cells[i, 2].value = id1;
                                                    worksheetPlansatirada.Rows.Cells[i, 3].value = reader2["ime"].ToString();
                                                    worksheetPlansatirada.Rows.Cells[i, 4].value = reader2["sifrarm"].ToString();
                                                    worksheetPlansatirada.Rows.Cells[i, 5].value = reader2["mt"].ToString();
                                                    worksheetPlansatirada.Rows.Cells[i, 6].value = ozn1;
                                                    worksheetPlansatirada.Rows.Cells[i, 7].value = (ukminuta1 / 60).ToString();
                                                    //worksheetPlansatirada.Rows.Cells[i, 8].value = reader2["RadnoMjesto"].ToString();
                                                    worksheetPlansatirada.Rows.Cells[i, 11].value = "11.provjeriti >> ";


                                                }
                                                else
                                                {
                                                    using (SqlConnection cn20 = new SqlConnection(connectionStringRFIND))
                                                    {
                                                        ozn1 = ozn1.Replace("0e;", "");

                                                        cn20.Open();
                                                        string sql20 = "update " + baza1 + " set " + danxx + "= '" + ozn1 + "' where  mjesec=" + mjes1 + " and godina=" + god1 + " and radnikid= " + id1 + " and firma='" + firma1 + "'";
                                                        SqlCommand sqlc20 = new SqlCommand(sql20, cn20);
                                                        SqlDataReader reader20 = sqlc20.ExecuteReader();
                                                        cn20.Close();

                                                    }
                                                }
                                            }
                                            else if (jj > 0)
                                            {
                                                string napomena1 = reader["napomena1"].ToString();
                                                if (napomena1 == "2")
                                                {
                                                    Console.WriteLine("2.Potrebna provjera za dan " + danxx + " Ime=" + radnikk + "   id1 = " + id1 + " Firma = " + firma1 + " ozn1 " + ozn1 + " odradio ukupno minuta " + ukminuta1.ToString() + "-- pomoćni radnik ?");
                                                    //Console.ReadKey();
                                                    i++;
                                                    worksheetPlansatirada.Rows.Cells[i, 1].value = firma1;
                                                    worksheetPlansatirada.Rows.Cells[i, 2].value = id1;
                                                    worksheetPlansatirada.Rows.Cells[i, 3].value = radnikk;
                                                    worksheetPlansatirada.Rows.Cells[i, 4].value = reader2["sifrarm"].ToString();
                                                    worksheetPlansatirada.Rows.Cells[i, 5].value = reader2["mt"].ToString();
                                                    worksheetPlansatirada.Rows.Cells[i, 6].value = ozn1;
                                                    worksheetPlansatirada.Rows.Cells[i, 7].value = (ukminuta1 / 60).ToString();
                                                    // worksheetPlansatirada.Rows.Cells[i, 8].value = reader["RadnoMjesto"].ToString();
                                                    worksheetPlansatirada.Rows.Cells[i, 11].value = "12.provjeriti >> ";



                                                }

                                            }

                                        }


                                    }
                                }



                            }



                        }




                    }

                }

                //return;
                // 2. korak oni koji nisu upisani u prvom koraku, i imaju više od 90 minuta

                sql1 = "select r.neradi,p.firma,p.radnikid,p.ime," + danxx + ",v.dosao,v.otisao,v.ukupno_minuta,v.smjena,a.fixnaisplata,a.mt from " + baza1 + " p left join feroapp.dbo.radnici r on r.id_fink = p.radnikid and r.id_firme = p.Firma " +
                "left join rfind.dbo.radnici_ a on a.id_radnika = r.id_radnika left join rfind.dbo.pregledvremena v on v.idradnika = a.id and v.Datum = '" + djucer + "' where mjesec = " + mjes1 + " and godina = " + god1 + " and v.ukupno_minuta > 90 and " + danxx + " is null  order by ime";
                //                "left join rfind.dbo.radnici_ a on a.id_radnika = r.id_radnika left join rfind.dbo.pregledvremena v on v.idradnika = a.id and v.Datum = '" + djucer + "' where mjesec = " + mjes1 + " and godina = " + god1 + " and v.ukupno_minuta > 465 and a.mt in (700,701,712, 716,704,710,702,713) and " + danxx + " is null  order by ime";

                using (SqlConnection cn = new SqlConnection(connectionStringRFIND))
                {

                    cn.Open();
                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);
                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string id1, firma1, sql2, ozn0, ozn1, firma10 = "", id10 = "", smjena1 = "", o1 = "", fi1 = "";
                    DateTime dosao;
                    int prviputa = 1, minuta, s1, rv, mt1;


                    while (reader.Read())
                    {
                        if (reader["ime"] != DBNull.Value)
                        {

                            if (reader[danxx] == DBNull.Value)
                            {// ako je null

                                ozn1 = "0e";
                                id1 = reader["radnikid"].ToString();
                                smjena1 = reader["smjena"].ToString().TrimEnd();
                                minuta = (int.Parse)(reader["ukupno_minuta"].ToString());
                                dosao = (DateTime)(reader["dosao"]);
                                int h1 = dosao.Hour;
                                int min1 = dosao.Minute;
                                firma1 = reader["firma"].ToString();
                                mt1 = (int.Parse)(reader["mt"].ToString());
                                fi1 = reader["fixnaisplata"].ToString();
                                rv = 7;

                                if (dayofweek1 > 0 && dayofweek1 < 6)  // od ponedjeljka do petka
                                {
                                    rv = 7;
                                }
                                else if (dayofweek1 == 6 && (smjena1 != "3"))  // subota 1 i 2 smjena
                                {
                                    rv = 5;
                                }
                                else // ostalo preskoči
                                {
                                    if (mt1 == 702)
                                    {
                                        if (dayofweek1 == 6 || dayofweek1 == 0)
                                        {
                                            rv = 5;
                                        }
                                    }
                                    else
                                    {
                                        continue;
                                    }
                                }
                                if (id1 == "172")
                                {
                                    int hh = 0;
                                }


                                s1 = minuta / 60;
                                int razlika = minuta - s1 * 60;
                                if (razlika > 45)
                                {
                                    s1 = s1 + 1;
                                }

                                // ako šteler dođe prerano, Oršiček Nenad
                                if (mt1 == 716 && ((h1 == 13) || (h1 == 21) || (h1 == 5)) && (min1 < 20))
                                {
                                    s1 = s1 - 1;
                                }

                                int o1broj = (s1 - rv);
                                o1 = o1broj.ToString();
                                if (fi1 == "1")    // ako ima fiksnu isplatu onda 0j
                                {
                                    o1 = "0";
                                }

                                if ((smjena1 == "3" && dayofweek1 == 6) || (dayofweek1 == 0))
                                {

                                    if (mt1 == 702)
                                    { }
                                    else
                                    {
                                        continue;
                                    }
                                }

                                if (smjena1 == "3")
                                {
                                    ozn1 = "n";
                                }
                                if (smjena1 == "2")
                                {
                                    ozn1 = "p";
                                }
                                if (smjena1 == "1")
                                {
                                    ozn1 = "j";
                                }
                                // ako je odradio više od 8 sati zaredom , nije 2 j nego 1j:-6p= 9 sati

                                if (((int.Parse)(o1) > 1) && (ozn1 == "j"))
                                {
                                    if (rv == 7)     // ponedj...petak
                                    {
                                        ozn1 = "1j:" + (o1broj - rv - 1).ToString() + "p";
                                    }
                                    else if (((int.Parse)(o1) > 3) && (ozn1 == "j") && (rv == 5))   // subota
                                    {
                                        ozn1 = "3j:" + (o1broj - rv - 2).ToString() + "p";
                                    }
                                    else  // 3j
                                    {
                                        ozn1 = o1 + ozn1;
                                    }

                                }
                                else
                                {
                                    ozn1 = o1 + ozn1;
                                }

                                using (SqlConnection cn20 = new SqlConnection(connectionStringRFIND))
                                {
                                    cn20.Open();
                                    if (id1 == "1465")
                                    {
                                        int hh = 0;
                                    }
                                    string sql20 = "update " + baza1 + " set " + danxx + "= '" + ozn1 + "' where  mjesec=" + mjes1 + " and godina=" + god1 + " and radnikid= " + id1 + " and firma='" + firma1 + "'";
                                    SqlCommand sqlc20 = new SqlCommand(sql20, cn20);
                                    SqlDataReader reader20 = sqlc20.ExecuteReader();
                                    cn20.Close();
                                }


                            }



                        }
                    }

                }

                //return;
                // 3.korak -- plansatirada 2 korak, oni koji upisano prethodno ( radnici,šteleri oni koji su zaplanirani i imaju >450 minuta)

                // Režija, svima koji se pojave ide 0j
                sql1 = "select r.neradi,p.firma,p.*,a.neradi from " + baza1 + " p left join feroapp.dbo.radnici r on r.id_fink = p.radnikid and r.id_firme = p.Firma left join rfind.dbo.radnici_ a on a.id_radnika = r.id_radnika left join rfind.dbo.pregledvremena v on v.IDRadnika=a.id and v.datum= '" + djucer + "' where a.neradi=0 and mjesec = " + mjes1 + " and godina = " + god1 + " and a.datumzaposlenja<='" + djucer + "' and r.sifrarm='Režija' and year(v.dosao)!=1900 and " + danxx + " is null order by ime ";
                //                sql1 = "select r.neradi,p.firma,p.*,a.neradi from " + baza1 + " p left join feroapp.dbo.radnici r on r.id_fink = p.radnikid and r.id_firme = p.Firma left join rfind.dbo.radnici_ a on a.id_radnika = r.id_radnika left join rfind.dbo.pregledvremena v on v.IDRadnika=a.id and v.datum= '" + djucer + "' where a.neradi=0 and mjesec = " + mjes1 + " and godina = " + god1 + " and a.datumzaposlenja<='" + djucer + "' and ( r.sifrarm='Režija' or (mt=707 ) ) and v.radnomjesto not in ('B.O.','G.O.')  and " + danxx + " is null order by ime ";
                //sql1 = "select r.neradi,p.firma,p.*,dan17,a.neradi from " + baza1 + " p left join feroapp.dbo.radnici r on r.id_fink = p.radnikid and r.id_firme = p.Firma left join rfind.dbo.radnici_ a on a.id_radnika = r.id_radnika where a.neradi=0 and mjesec = " + mjes1 + " and godina = " + god1 + " and a.mt in (700,701,704,713,712, 716) and  a.datumzaposlenja<='" + djucer + "' and " + danxx + " is null order by ime ";

                using (SqlConnection cn = new SqlConnection(connectionStringRFIND))
                {

                    cn.Open();
                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);
                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string id1, firma1, sql2, ozn0, ozn1, firma10 = "", id10 = "", smjena1 = "";
                    int prviputa = 1;

                    while (reader.Read())
                    {
                        if (reader["ime"] != DBNull.Value)
                        {

                            if (reader[danxx] == DBNull.Value)
                            {// ako je null

                                ozn1 = "0j";
                                id1 = reader["radnikid"].ToString();
                                firma1 = reader["firma"].ToString();
                                if ((dayofweek1 == 0))
                                {
                                    continue;
                                }

                                using (SqlConnection cn20 = new SqlConnection(connectionStringRFIND))
                                {
                                    cn20.Open();
                                    string sql20 = "update " + baza1 + " set " + danxx + "= '" + ozn1 + "' where  mjesec=" + mjes1 + " and godina=" + god1 + " and radnikid= " + id1 + " and firma='" + firma1 + "'";
                                    SqlCommand sqlc20 = new SqlCommand(sql20, cn20);
                                    SqlDataReader reader20 = sqlc20.ExecuteReader();
                                    cn20.Close();
                                }

                            }

                        }
                    }

                }


                // puni 0e
                sql1 = "select r.neradi,p.firma,p.*,a.neradi from " + baza1 + " p left join feroapp.dbo.radnici r on r.id_fink = p.radnikid and r.id_firme = p.Firma left join rfind.dbo.radnici_ a on a.id_radnika = r.id_radnika left join pregledvremena v on v.idradnika=a.id and v.datum='" + djucer + "' where a.neradi=0 and mjesec = " + mjes1 + " and godina = " + god1 + " and a.datumzaposlenja<='" + djucer + "' and " + danxx + " is null and charindex('4. SMJENA',v.radnomjesto)=0  order by ime ";
                //sql1 = "select r.neradi,p.firma,p.*,dan17,a.neradi from " + baza1 + " p left join feroapp.dbo.radnici r on r.id_fink = p.radnikid and r.id_firme = p.Firma left join rfind.dbo.radnici_ a on a.id_radnika = r.id_radnika where a.neradi=0 and mjesec = " + mjes1 + " and godina = " + god1 + " and a.mt in (700,701,704,713,712, 716) and  a.datumzaposlenja<='" + djucer + "' and " + danxx + " is null order by ime ";

                using (SqlConnection cn = new SqlConnection(connectionStringRFIND))
                {

                    cn.Open();
                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);
                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string id1, firma1, sql2, ozn0, ozn1, firma10 = "", id10 = "", smjena1 = "";
                    int prviputa = 1;

                    while (reader.Read())
                    {
                        if (reader["ime"] != DBNull.Value)
                        {

                            if (reader[danxx] == DBNull.Value)
                            {// ako je null

                                ozn1 = "0e";
                                id1 = reader["radnikid"].ToString();
                                firma1 = reader["firma"].ToString();
                                var list = new List<int> { 128, 971, 257, 880 };   // Horvatic, Legac H,Z,B
                                if (id1 == "971")
                                {
                                    int z = 0;
                                }
                                var intVar = (int.Parse)(id1);
                                var exists = list.Contains(intVar);
                                if (exists && firma1 == "1")
                                {
                                    ozn1 = "0j";
                                }
                                list = new List<int> { 4, 71 };  // Branimir,Lana - Tokabu
                                intVar = (int.Parse)(id1);
                                exists = list.Contains(intVar);
                                if (exists && firma1 == "3")
                                {
                                    ozn1 = "0j";
                                }

                                // preskači subotu i nedjelju
                                if ((dayofweek1 == 6) || (dayofweek1 == 0))
                                {
                                    continue;
                                }

                                using (SqlConnection cn20 = new SqlConnection(connectionStringRFIND))
                                {
                                    cn20.Open();
                                    string sql20 = "update " + baza1 + " set " + danxx + "= '" + ozn1 + "' where  mjesec=" + mjes1 + " and godina=" + god1 + " and radnikid= " + id1 + " and firma='" + firma1 + "'";
                                    SqlCommand sqlc20 = new SqlCommand(sql20, cn20);
                                    SqlDataReader reader20 = sqlc20.ExecuteReader();
                                    cn20.Close();
                                }

                            }

                        }
                    }

                }


                // update za upravu, ako je nešto upisano ili vikend preskoči
                if ((dayofweek1 == 6) || (dayofweek1 == 0))
                {
                    // ako je nedjelja ili subota
                }
                else
                {
                    using (SqlConnection cn20 = new SqlConnection(connectionStringRFIND))
                    {
                        cn20.Open();
                        string sql20 = "update " + baza1 + " set " + danxx + "='0j' where  mjesec=" + mjes1 + " and godina=" + god1 + " and radnikid in ( 128,971,257,880) and firma=1 and " + danxx + " is null ";
                        SqlCommand sqlc20 = new SqlCommand(sql20, cn20);
                        SqlDataReader reader20 = sqlc20.ExecuteReader();
                        cn20.Close();

                        cn20.Open();
                        sql20 = "update " + baza1 + " set " + danxx + "='0j' where  mjesec=" + mjes1 + " and godina=" + god1 + " and radnikid in (4,71)  and firma=3 and " + danxx + " is null";
                        sqlc20 = new SqlCommand(sql20, cn20);
                        reader20 = sqlc20.ExecuteReader();
                        cn20.Close();

                    }
                }

                // puni 4. sheet  za one koji imaju 0e ili null
                // preskoči tip=3
                if (dayofweek1 == 9999)  // za nedjelju ne puni 4.sheet 
                {
                    sql1 = "rfind.dbo.izostanci3 '" + dat1 + "',3";
                    using (SqlConnection cn = new SqlConnection(connectionStringRFIND))
                    {

                        cn.Open();
                        SqlCommand sqlCommand = new SqlCommand(sql1, cn);
                        //i = 1;

                        worksheetPlansatirada.Rows.Cells[i, 6].value = danxx;  // u zaglavlje stavi pravu oznaku dana
                        string ozn1;
                        SqlDataReader reader = sqlCommand.ExecuteReader();
                        string kupac, vrstap, ii, sbsati, napom1 = "", id1;
                        int l1 = 0, bs = 0, bsat = 0;

                        while (reader.Read())
                        {
                            if (reader["firma"] != DBNull.Value)
                            {
                                id1 = reader["radnikid"].ToString();
                                if (id1 == "1659")
                                {
                                    int z23 = 0;
                                }
                                ozn1 = reader[danxx].ToString();
                                napom1 = "";

                                if (reader["Tip"].ToString() == "1")
                                {
                                    sbsati = reader["Sati"].ToString();
                                    l1 = ozn1.Length;

                                    if (sbsati == "")
                                    {
                                        bsat = 0;
                                    }
                                    else
                                    {
                                        bsat = (int.Parse)(reader["Sati"].ToString());
                                    }


                                    if (((ozn1 == "0e") || (ozn1 == "")) && (bsat == 0))  // ako je već upisano 0e i ima 0 sati
                                    {

                                    }
                                    else
                                    {

                                        if (ozn1.Contains(":"))
                                        {
                                            string prvidio, drugidio;
                                            int bs1 = 0, bs2 = 0;
                                            if (ozn1.Contains("j") || ozn1.Contains("p") || ozn1.Contains("n"))
                                            {

                                                prvidio = ozn1.Substring(0, ozn1.IndexOf(':'));
                                                drugidio = ozn1.Substring(ozn1.IndexOf(':') + 1);
                                                ozn1 = prvidio;
                                                if (ozn1.Contains("j") || ozn1.Contains("p") || ozn1.Contains("n"))
                                                {
                                                    l1 = ozn1.Length;
                                                    ii = ozn1.Substring(2);
                                                    if (l1 == 2)
                                                    {
                                                        bs1 = (int.Parse)(ozn1.Substring(0, 1)) + brojsati;
                                                    }
                                                    else if (l1 == 3)
                                                    {
                                                        bs1 = (int.Parse)(ozn1.Substring(0, 2)) + brojsati;
                                                    }
                                                }
                                                ozn1 = drugidio;
                                                if (ozn1.Contains("j") || ozn1.Contains("p") || ozn1.Contains("n"))
                                                {
                                                    l1 = ozn1.Length;
                                                    ii = ozn1.Substring(2);
                                                    if (l1 == 2)
                                                    {
                                                        bs2 = (int.Parse)(ozn1.Substring(0, 1)) + brojsati;
                                                    }
                                                    else if (l1 == 3)
                                                    {
                                                        bs2 = (int.Parse)(ozn1.Substring(0, 2)) + brojsati;
                                                    }
                                                }
                                                bs = bs1 + bs2;

                                                if (bs == bsat)
                                                {
                                                    napom1 = "Ok";
                                                }
                                                if ((bs - bsat) > 1)
                                                {
                                                    napom1 = "Provjeriti sate, upisano više nego po kartici ! >> ";
                                                }
                                                else if ((bsat - bs) > 1)
                                                {
                                                    napom1 = "Provjeriti sate, upisano manje nego po kartici ! >> ";
                                                }


                                                //     ozn1 = ozn1.Remove('j');

                                            }
                                        }
                                        else  // ako nema :
                                        {
                                            if (ozn1.Contains("j") || ozn1.Contains("p") || ozn1.Contains("n"))
                                            {
                                                l1 = ozn1.Length;
                                                ii = ozn1.Substring(2);
                                                if (l1 == 2)
                                                {
                                                    bs = (int.Parse)(ozn1.Substring(0, 1)) + brojsati;
                                                }
                                                else if (l1 == 3)
                                                {
                                                    bs = (int.Parse)(ozn1.Substring(0, 2)) + brojsati;
                                                }
                                                if (bs == bsat)
                                                {
                                                    napom1 = "Ok";
                                                }
                                                if (sbsati == "")
                                                {
                                                    napom1 = "Nije zaplaniran ! >>";
                                                }
                                                if (sbsati == "0")
                                                {
                                                    napom1 = "Provjeriti dali se dobro registrirao ! >> ";
                                                }
                                                if ((bs - bsat) > 1)
                                                {
                                                    napom1 = "Provjeriti sate, upisano više nego po kartici ! >> ";
                                                }
                                                else if ((bsat - bs) > 1)
                                                {
                                                    napom1 = "Provjeriti sate, upisano manje nego po kartici ! >> ";
                                                }


                                            }
                                        }

                                    }
                                }

                                i++;
                                worksheetPlansatirada.Rows.Cells[i, 1].value = reader["firma"].ToString();
                                worksheetPlansatirada.Rows.Cells[i, 2].value = reader["radnikid"].ToString();
                                worksheetPlansatirada.Rows.Cells[i, 3].value = reader["ime"].ToString();
                                worksheetPlansatirada.Rows.Cells[i, 4].value = reader["sifrarm"].ToString();
                                worksheetPlansatirada.Rows.Cells[i, 5].value = reader["mt"].ToString();
                                worksheetPlansatirada.Rows.Cells[i, 6].value = reader[danxx].ToString();
                                worksheetPlansatirada.Rows.Cells[i, 7].value = reader["Sati"].ToString();
                                worksheetPlansatirada.Rows.Cells[i, 8].value = reader["RadnoMjesto"].ToString();
                                worksheetPlansatirada.Rows.Cells[i, 9].value = reader["Hala"].ToString();
                                worksheetPlansatirada.Rows.Cells[i, 10].value = reader["Smjena"].ToString();
                                worksheetPlansatirada.Rows.Cells[i, 11].value = napom1 + ' ' + reader["Napomena"].ToString();
                                worksheetPlansatirada.Rows.Cells[i, 12].value = reader["Tip"].ToString();


                                if (((ozn1 == "0e") || (ozn1 == "")) && (bsat > 0))  // ako je već upisano 0e i ima 0 sati
                                {
                                    worksheetPlansatirada.Rows.Cells[i, 11].value = "Upisano 0e a došao na posao !! >>" + reader["Napomena"].ToString();
                                }

                            }
                        }
                    }
                }
                //  return;
                // ako ima upisano 0e i dosao je 
                sql1 = "rfind.dbo.izostanci3 '" + dat1 + "',4";
                using (SqlConnection cn = new SqlConnection(connectionStringRFIND))
                {

                    cn.Open();
                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);
                    //i = 1;

                    string ozn1;
                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string kupac, vrstap, ii, sbsati, napom1 = "";
                    int l1 = 0, bs = 0, bsat = 0;

                    while (reader.Read())
                    {

                        i++;
                        worksheetPlansatirada.Rows.Cells[i, 1].value = reader["firma"].ToString();
                        worksheetPlansatirada.Rows.Cells[i, 2].value = reader["radnikid"].ToString();
                        worksheetPlansatirada.Rows.Cells[i, 3].value = reader["ime"].ToString();
                        worksheetPlansatirada.Rows.Cells[i, 4].value = reader["sifrarm"].ToString();
                        worksheetPlansatirada.Rows.Cells[i, 5].value = reader["mt"].ToString();
                        worksheetPlansatirada.Rows.Cells[i, 6].value = reader["datum"].ToString();
                        worksheetPlansatirada.Rows.Cells[i, 7].value = reader["Sati"].ToString();
                        worksheetPlansatirada.Rows.Cells[i, 8].value = reader["RadnoMjesto"].ToString();
                        worksheetPlansatirada.Rows.Cells[i, 9].value = reader["Hala"].ToString();
                        worksheetPlansatirada.Rows.Cells[i, 10].value = reader["Smjena"].ToString();
                        worksheetPlansatirada.Rows.Cells[i, 11].value = "Ima upisano 0e ili null a došao --" + reader["napomena"].ToString();
                        worksheetPlansatirada.Rows.Cells[i, 12].value = reader["Tip"].ToString();



                    }
                }
                // nezaplanirani
                sql1 = "rfind.dbo.izostanci3 '" + dat1 + "',5";
                using (SqlConnection cn = new SqlConnection(connectionStringRFIND))
                {

                    cn.Open();
                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);
                    //i = 1;

                    string ozn1;
                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string kupac, vrstap, ii, sbsati, napom1 = "";
                    int l1 = 0, bs = 0, bsat = 0;

                    while (reader.Read())
                    {

                        i++;

                        worksheetPlansatirada.Rows.Cells[i, 1].value = reader["firma"].ToString();
                        worksheetPlansatirada.Rows.Cells[i, 2].value = reader["radnikid"].ToString();
                        worksheetPlansatirada.Rows.Cells[i, 3].value = reader["ime"].ToString();
                        worksheetPlansatirada.Rows.Cells[i, 4].value = reader["sifrarm"].ToString();
                        worksheetPlansatirada.Rows.Cells[i, 5].value = reader["mt"].ToString();
                        worksheetPlansatirada.Rows.Cells[i, 6].value = reader["ozn1"].ToString();
                        worksheetPlansatirada.Rows.Cells[i, 7].value = reader["Sati"].ToString();
                        worksheetPlansatirada.Rows.Cells[i, 8].value = reader["RadnoMjesto"].ToString();
                        worksheetPlansatirada.Rows.Cells[i, 9].value = reader["Hala"].ToString();
                        worksheetPlansatirada.Rows.Cells[i, 10].value = reader["Smjena"].ToString();
                        worksheetPlansatirada.Rows.Cells[i, 11].value = "Nezaplanirani --" + reader["napomena"].ToString();
                        worksheetPlansatirada.Rows.Cells[i, 12].value = reader["Tip"].ToString();

                    }
                }

                // provjera sati
                sql1 = "rfind.dbo.izostanci3 '" + dat1 + "',6";
                using (SqlConnection cn = new SqlConnection(connectionStringRFIND))
                {

                    cn.Open();
                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);
                    //i = 1;

                    string ozn1;
                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string kupac, vrstap, ii, sbsati, napom1 = "";
                    int l1 = 0, bs = 0, bsat = 0;

                    while (reader.Read())
                    {

                        ozn1 = reader["ozn1"].ToString();
                        sbsati = reader["Sati"].ToString();
                        bsat = (int.Parse)(sbsati);
                        napom1 = "";
                        if (ozn1.Contains(":"))
                        {
                            string prvidio, drugidio;
                            int bs1 = 0, bs2 = 0;
                            if (ozn1.Contains("j") || ozn1.Contains("p") || ozn1.Contains("n"))
                            {

                                prvidio = ozn1.Substring(0, ozn1.IndexOf(':'));
                                drugidio = ozn1.Substring(ozn1.IndexOf(':') + 1);
                                ozn1 = prvidio;
                                if (ozn1.Contains("j") || ozn1.Contains("p") || ozn1.Contains("n"))
                                {
                                    l1 = ozn1.Length;
                                    ii = ozn1.Substring(2);
                                    if (l1 == 2)
                                    {
                                        bs1 = (int.Parse)(ozn1.Substring(0, 1)) + brojsati;
                                    }
                                    else if (l1 == 3)
                                    {
                                        bs1 = (int.Parse)(ozn1.Substring(0, 2)) + brojsati;
                                    }
                                }
                                ozn1 = drugidio;
                                if (ozn1.Contains("j") || ozn1.Contains("p") || ozn1.Contains("n"))
                                {
                                    l1 = ozn1.Length;
                                    ii = ozn1.Substring(2);
                                    if (l1 == 2)
                                    {
                                        bs2 = (int.Parse)(ozn1.Substring(0, 1)) + brojsati;
                                    }
                                    else if (l1 == 3)
                                    {
                                        bs2 = (int.Parse)(ozn1.Substring(0, 2)) + brojsati;
                                    }
                                }
                                bs = bs1 + bs2;

                                if (bs == bsat)
                                {
                                    napom1 = "Ok";
                                }
                                if ((bs - bsat) > 1)
                                {
                                    napom1 = "Provjeriti sate, upisano više nego po kartici ! >> ";
                                }
                                else if ((bsat - bs) > 1)
                                {
                                    napom1 = "Provjeriti sate, upisano manje nego po kartici ! >> ";
                                }

                                //     ozn1 = ozn1.Remove('j');

                            }
                        }
                        else  // ako nema :
                        {
                            if (ozn1.Contains("j") || ozn1.Contains("p") || ozn1.Contains("n"))
                            {
                                l1 = ozn1.Length;
                                ii = ozn1.Substring(2);
                                if (l1 == 2)
                                {
                                    bs = (int.Parse)(ozn1.Substring(0, 1)) + brojsati;
                                }
                                else if (l1 == 3)
                                {
                                    bs = (int.Parse)(ozn1.Substring(0, 2)) + brojsati;
                                }
                                if (bs == bsat)
                                {
                                    napom1 = "Ok";
                                }
                                if (sbsati == "")
                                {
                                    napom1 = "Nije zaplaniran ! >>";
                                }
                                if (sbsati == "0")
                                {
                                    napom1 = "Provjeriti dali se dobro registrirao ! >> ";
                                }

                                if (bsat > 0 && ozn1 == "0j")
                                {
                                }
                                else
                                {
                                    if ((bs - bsat) > 1)
                                    {
                                        napom1 = "Provjeriti sate, upisano više nego po kartici ! >> ";
                                    }
                                    else if ((bsat - bs) > 1)
                                    {
                                        napom1 = "Provjeriti sate, upisano manje nego po kartici ! >> ";
                                    }
                                }

                            }
                            if (((ozn1 == "0e") || (ozn1 == "")) && bsat == 0)
                            {
                                napom1 = "Nije došao ima upisano 0e ili null, ok ";
                            }
                            else if (((ozn1 == "0e") || (ozn1 == "")) && bsat != 0)
                            {
                                napom1 = "Ima upisano 0e i bio je " + sbsati + " sati ";
                            }


                        }
                        if (ozn1=="7g" && bsat==0)
                        {
                            napom1 = "Ok";
                        }
                        string mt1 = reader["mt"].ToString();
                        string smjena1 = reader["Smjena"].ToString();

                        if ( ((smjena1 == "3" && dayofweek1 == 6) || (dayofweek1 == 0)) && (!mt1.Contains("702")) )
                        {
                            continue;
                        }
                                               

                        if (napom1.Contains("Ok") || (ozn1 == "0j"))
                        { }
                        else
                        {
                            i++;
                            worksheetPlansatirada.Rows.Cells[i, 1].value = reader["firma"].ToString();
                            worksheetPlansatirada.Rows.Cells[i, 2].value = reader["radnikid"].ToString();
                            worksheetPlansatirada.Rows.Cells[i, 3].value = reader["ime"].ToString();
                            worksheetPlansatirada.Rows.Cells[i, 4].value = reader["sifrarm"].ToString();

                            worksheetPlansatirada.Rows.Cells[i, 5].value = reader["mt"].ToString();
                            worksheetPlansatirada.Rows.Cells[i, 6].value = reader["ozn1"].ToString();
                            worksheetPlansatirada.Rows.Cells[i, 7].value = reader["Sati"].ToString();
                            worksheetPlansatirada.Rows.Cells[i, 8].value = reader["RadnoMjesto"].ToString();
                            worksheetPlansatirada.Rows.Cells[i, 9].value = reader["Hala"].ToString();
                            worksheetPlansatirada.Rows.Cells[i, 10].value = reader["Smjena"].ToString();
                            worksheetPlansatirada.Rows.Cells[i, 11].value = napom1 + " --- " + reader["napomena"].ToString();
                            worksheetPlansatirada.Rows.Cells[i, 12].value = reader["Tip"].ToString();

                        }

                    }
                }

                //  return;


            }

                //end 1


                // end plansatirada

                // izvještaj za plansati rada 0e


                for (int ws = 1; ws < 2; ws++)
            {                

                sql1 = "rfind.dbo.izostanci '" + dat1 + "',1";
                using (SqlConnection cn = new SqlConnection(connectionStringRFIND))
                {
                    
                    cn.Open();
                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);
                    i = 1;
                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string kupac, vrstap;

                    while (reader.Read())
                    {
                        if (reader["datum"] != DBNull.Value)
                        {
                            i++;

                            worksheetIZO.Rows.Cells[i, 1].value = reader["datum"].ToString();
                            worksheetIZO.Rows.Cells[i, 2].value = reader["id"].ToString();
                            worksheetIZO.Rows.Cells[i, 3].value = reader["ime"].ToString();
                            worksheetIZO.Rows.Cells[i, 4].value = reader["prezime"].ToString();
                            worksheetIZO.Rows.Cells[i, 5].value = reader["mt_naziv"].ToString();
                            worksheetIZO.Rows.Cells[i, 6].value = reader["mt"].ToString();
                            worksheetIZO.Rows.Cells[i, 7].value = reader["radnomjesto"].ToString();
                            worksheetIZO.Rows.Cells[i, 8].value = reader["hala"].ToString();
                            worksheetIZO.Rows.Cells[i, 9].value = reader["smjena"].ToString();
                            worksheetIZO.Rows.Cells[i, 10].value = reader["napomena"].ToString();
                            worksheetIZO.Rows.Cells[i, 11].value = reader["napomena2"].ToString();
                            worksheetIZO.Rows.Cells[i, 12].value = reader["dnzp"].ToString();
                            worksheetIZO.Rows.Cells[i, 13].value = reader["kasni"].ToString();
                            worksheetIZO.Rows.Cells[i, 14].value = reader["nedostaje"].ToString();
                            worksheetIZO.Rows.Cells[i, 15].value = reader["preranootišao"].ToString();

                        }

                    }                

                }

                sql1 = "rfind.dbo.izostanci2 '" + dat1 + "',1";

                Worksheet worksheetNP = workbook.Worksheets.Item[2] as Worksheet;

                using (SqlConnection cn = new SqlConnection(connectionStringRFIND))
                {

                    cn.Open();
                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string kupac, vrstap;
                    
                    while (reader.Read())
                    {
                        if (reader["id"] != DBNull.Value)
                        {
                            j++;

                            worksheetNP.Rows.Cells[j, 1].value = reader["ID"].ToString();                                                                                                                                                                                       
                            worksheetNP.Cells[j, 2].value = reader["ime"].ToString();
                            worksheetNP.Rows.Cells[j, 3].value = reader["prezime"].ToString();
                            worksheetNP.Rows.Cells[j, 4].value = reader["vrijeme"].ToString();
                            worksheetNP.Rows.Cells[j, 5].value = reader["name"].ToString();

                        }

                    }
                }     // end for ws<=4

                sql1 = "rfind.dbo.izostanci3 '" + dat1 + "',1";

                Worksheet worksheetDUPL = workbook.Worksheets.Item[3] as Worksheet;

                using (SqlConnection cn = new SqlConnection(connectionStringRFIND))
                {

                    cn.Open();
                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);
                    k = 1;
                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string kupac, vrstap;                    
                    while (reader.Read())
                    {
                        if (reader["id"] != DBNull.Value)
                        {
                            k++;

                            worksheetDUPL.Rows.Cells[k, 1].value = reader["ID"].ToString();
                            worksheetDUPL.Cells[k, 2].value = reader["prezime"].ToString();
                            worksheetDUPL.Rows.Cells[k, 3].value = reader["ime"].ToString();
                            worksheetDUPL.Rows.Cells[k, 4].value = reader["rbroj"].ToString();
                            worksheetDUPL.Rows.Cells[k, 5].value = reader["datum"].ToString();
                            worksheetDUPL.Rows.Cells[k, 6].value = reader["hala"].ToString();
                            worksheetDUPL.Rows.Cells[k, 7].value = reader["smjena"].ToString();
                            worksheetDUPL.Rows.Cells[k, 8].value = reader["radnomjesto"].ToString();

                        }

                    }
                }     // end for ws<=4

                if (ws == 1)   // problematične prijave
                {
                    sql1 = "rfind.dbo.izostanci3 '" + dat1 + "',2";

                    Worksheet worksheetrfid = workbook.Worksheets.Item[4] as Worksheet;

                    using (SqlConnection cn = new SqlConnection(connectionStringRFIND))
                    {

                        cn.Open();
                        SqlCommand sqlCommand = new SqlCommand(sql1, cn);
                        k = 1;
                        SqlDataReader reader = sqlCommand.ExecuteReader();
                        string kupac, vrstap;
                        while (reader.Read())
                        {
                            if (reader["id"] != DBNull.Value)
                            {
                                k++;

                                worksheetrfid.Rows.Cells[k, 1].value = reader["ID"].ToString();
                                worksheetrfid.Cells[k, 2].value = reader["prezime"].ToString();
                                worksheetrfid.Rows.Cells[k, 3].value = reader["ime"].ToString();
                                worksheetrfid.Rows.Cells[k, 4].value = reader["rbroj"].ToString();
                                worksheetrfid.Rows.Cells[k, 5].value = reader["datum"].ToString();
                                worksheetrfid.Rows.Cells[k, 6].value = reader["hala"].ToString();
                                worksheetrfid.Rows.Cells[k, 7].value = reader["smjena"].ToString();
                                worksheetrfid.Rows.Cells[k, 8].value = reader["napomena"].ToString();
                                worksheetrfid.Rows.Cells[k, 9].value = reader["dosao"].ToString();
                                worksheetrfid.Rows.Cells[k, 10].value = reader["otisao"].ToString();
                                worksheetrfid.Rows.Cells[k, 11].value = reader["radnomjesto"].ToString();

                            }

                        }
                    }     // end for ws<=4
                }



                d2 = d2.AddDays(1);
                dat1 = d2.Year.ToString() + '-' + mm1 + d2.Month.ToString() + '-' + d2.Day.ToString();

            }

            DateTime jucer = DateTime.Now.AddDays(-1);
            
            sql1 = "rfind.dbo.PlanZadnjiDan " + jucer.Month.ToString() + ","+ jucer.Year.ToString();
            using (SqlConnection cn = new SqlConnection(connectionString))
            {
                cn.Open();
                SqlCommand sqlCommand = new SqlCommand(sql1, cn);
                k = 1;
       //         SqlDataReader reader = sqlCommand.ExecuteReader();
            }
            

            string smjenav;
            smjenav = "Smjena";
            Environment.SetEnvironmentVariable(smjenav, "3");

            //fileName = smj + "_" + fileName;
            var fi = new FileInfo(fileNameIzo);
            if (fi.Exists) File.Delete(fileNameIzo);
            
            excel.Application.ActiveWorkbook.SaveAs(fileNameIzo);                        

            Console.WriteLine("Snimljen file Izostanci  trenutno vrijeme " + DateTime.Now);
            //Console.ReadKey();
            //workbook.Close(false);
            //excel.Application.Quit();
            //excel.Quit();

            //workbook.Close(true, Type.Missing, Type.Missing);
            workbook.Close(false, Type.Missing, Type.Missing);
            excel.Application.Quit();
            excel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            workbook = null;
            app = null;

            //            string fileName = @"C:\brisi\dsr20062017.xlsm";
            MailMessage mail = new MailMessage("gasparic.s@feroimpex.hr", "gasparic.s@feroimpex.hr");

            if (test == 0)
                mail = new MailMessage("gasparic.s@feroimpex.hr", "cakanic.s@feroimpex.hr,gasparic.s@feroimpex.hr, srecckog@gmail.com,kicin.d@feroimpex.hr,hren.h@feroimpex.hr,husta.m@feroimpex.hr,deanovic.g@feroimpex.hr,jancic.d@feroimpex.hr,vladic.p@feroimpex.hr,francekovic.b@feroimpex.hr,hajtok.m@feroimpex.hr,biscan.s@feroimpex.hr,grgecic.d@feroimpex.hr,kolar.a@feroimpex.hr,steleri-p2@feroimpex.hr,steleri-p1@feroimpex.hr");

            SmtpClient client = new SmtpClient();
            client.Port = 25;
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            client.UseDefaultCredentials = false;
            client.Credentials = new System.Net.NetworkCredential("gasparic.s@feroimpex.hr", "gasparic1");

            client.Host = "mail.feroimpex.hr";
            //client.Host = "mail.feroimpex.hr";
            mail.Subject = "Izostanci " + datreps ;
            mail.Body = "Izostanci za prethodni i današnji dan " + datreps ;

            Attachment attachment = new Attachment(fileNameIzo, System.Net.Mime.MediaTypeNames.Application.Octet);
            System.Net.Mime.ContentDisposition disposition = attachment.ContentDisposition;
            disposition.CreationDate = File.GetCreationTime(fileNameIzo);
            disposition.ModificationDate = File.GetLastWriteTime(fileNameIzo);
            disposition.ReadDate = File.GetLastAccessTime(fileNameIzo);
            disposition.FileName = Path.GetFileName(fileNameIzo);
            disposition.Size = new FileInfo(fileNameIzo).Length;
            disposition.DispositionType = System.Net.Mime.DispositionTypeNames.Attachment;

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

        static private string provjerasati(string ozn1, string sbsati, int brojsati)
        {
            string prvidio, drugidio, ii, napom1 = "";
            int bs = 0, bs1 = 0, bs2 = 0, l1 = 0, bsat=0;
            
            l1 = ozn1.Length;

            if (sbsati == "")
            {
                bsat = 0;
            }
            else
            {
                bsat = (int.Parse)(sbsati);
            }

            if (ozn1.Contains(":"))
            {

                if (ozn1.Contains("j") || ozn1.Contains("p") || ozn1.Contains("n"))
                {

                    prvidio = ozn1.Substring(0, ozn1.IndexOf(':'));
                    drugidio = ozn1.Substring(ozn1.IndexOf(':') + 1);
                    ozn1 = prvidio;
                    if (ozn1.Contains("j") || ozn1.Contains("p") || ozn1.Contains("n"))
                    {
                        l1 = ozn1.Length;
                        ii = ozn1.Substring(2);
                        if (l1 == 2)
                        {
                            bs1 = (int.Parse)(ozn1.Substring(0, 1)) + brojsati;
                        }
                        else if (l1 == 3)
                        {
                            bs1 = (int.Parse)(ozn1.Substring(0, 2)) + brojsati;
                        }
                    }
                    ozn1 = drugidio;
                    if (ozn1.Contains("j") || ozn1.Contains("p") || ozn1.Contains("n"))
                    {
                        l1 = ozn1.Length;
                        ii = ozn1.Substring(2);
                        if (l1 == 2)
                        {
                            bs2 = (int.Parse)(ozn1.Substring(0, 1)) + brojsati;
                        }
                        else if (l1 == 3)
                        {
                            bs2 = (int.Parse)(ozn1.Substring(0, 2)) + brojsati;
                        }
                    }
                    bs = bs1 + bs2;

                    if (bs == bsat)
                    {
                        napom1 = "Ok";

                    }
                    if ((bs - bsat) > 1)
                    {
                        napom1 = "Provjeriti sate, upisano više nego po kartici ! >> ";
                    }
                    else if ((bsat - bs) > 1)
                    {
                        napom1 = "Provjeriti sate, upisano manje nego po kartici ! >> ";
                    }

                    //     ozn1 = ozn1.Remove('j');

                }

            }
            else  // ako nema :
            {
                if (ozn1.Contains("j") || ozn1.Contains("p") || ozn1.Contains("n"))
                {
                    l1 = ozn1.Length;
                    ii = ozn1.Substring(2);
                    if (l1 == 2)
                    {
                        bs = (int.Parse)(ozn1.Substring(0, 1)) + brojsati;
                    }
                    else if (l1 == 3)
                    {
                        bs = (int.Parse)(ozn1.Substring(0, 2)) + brojsati;
                    }
                    if (bs == bsat)
                    {
                        napom1 = "Ok";
                    }
                    if (sbsati == "")
                    {
                        napom1 = "Nije zaplaniran ! >>";
                    }
                    if (sbsati == "0")
                    {
                        napom1 = "Provjeriti dali se dobro registrirao ! >> ";
                    }

                    if (bsat > 0 && ozn1 == "0j")
                    {
                    }
                    else
                    {
                        if ((bs - bsat) > 1)
                        {
                            napom1 = "Provjeriti sate, upisano više nego po kartici ! >> ";
                        }
                        else if ((bsat - bs) > 1)
                        {
                            napom1 = "Provjeriti sate, upisano manje nego po kartici ! >> ";
                        }
                    }

                }
                if (((ozn1 == "0e") || (ozn1 == "")) && bsat == 0)
                {
                    napom1 = "Nije došao ima upisano 0e ili null, ok ";
                }
                else if (((ozn1 == "0e") || (ozn1 == "")) && bsat != 0)
                {
                    napom1 = "Ima upisano 0e i bio je " + sbsati + " sati ";
                }
                                             
            }
            return napom1;
        }
    }
}

