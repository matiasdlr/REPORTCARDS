using iTextSharp.text;
using iTextSharp.text.pdf;
using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Services;
using System.Web.UI;
using System.Web.UI.WebControls;




namespace HS_REPC_Final

{
   

    public partial class Default : System.Web.UI.Page
    {
        
        [WebMethod]
        public static string hrSelect(string gra)
        {
            string sql = string.Empty;
            string data = string.Empty;
            OracleConnection con = new OracleConnection();
            con.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings["conn"].ConnectionString;
            try
            {

                if (con.State == ConnectionState.Closed)
                {
                    con.Open();
                }

                //HR TEACHERS
                sql = "SELECT t.id,t.lastfirst teacher from sections sec";
                sql += " left join teachers t on sec.teacher = t.id";
                sql += " where sec.course_number like '%HR%' AND SEC.TERMID = 2800 AND GRADE_LEVEL ='" + gra + "'";
                sql += " order by teacher";
                
               
                OracleCommand cmd = new OracleCommand(sql, con);
                OracleDataReader odr = cmd.ExecuteReader();
                while (odr.Read())
                {
                    data += odr["id"].ToString() + '|';
                    data += odr["teacher"].ToString() + '^';

                }

                con.Close();
                con.Dispose();
            }
            catch (Exception ex)
            {
                throw;
            }
            return data;
        }

        
         [WebMethod]
        public static string hrgrade(string gra,int Htea)
        {
            string sql = string.Empty;
            string data = string.Empty;
            OracleConnection con = new OracleConnection();
            con.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings["conn"].ConnectionString;
            try
            {

                if (con.State == ConnectionState.Closed)
                {
                    con.Open();
                }


                //Hr students list
                sql = "SELECT S.STUDENT_NUMBER,S.FIRST_NAME||' '||S.LAST_NAME AS STUDENT,S.GRADE_LEVEL,T.FIRST_NAME||' '||T.LAST_NAME AS TEACHER FROM STUDENTS S";
                sql += " LEFT JOIN CC CO ON S.ID = CO.STUDENTID";
                sql += " LEFT JOIN TEACHERS T ON CO.TEACHERID = T.ID";
                sql += " LEFT JOIN COURSES C ON CO.COURSE_NUMBER = C.COURSE_NUMBER";
                sql += " WHERE S.ENROLL_STATUS = 0 AND C.COURSE_NAME LIKE '%HR%' AND CO.TERMID = 2800 AND S.GRADE_LEVEL = '" + gra + "' AND T.ID="+Htea+"";
                sql += " ORDER BY STUDENT";

                
                OracleCommand cmd = new OracleCommand(sql, con);
                OracleDataReader odr = cmd.ExecuteReader();
                while (odr.Read())
                {
                    data += odr["STUDENT_NUMBER"].ToString() + '|';
                    data += odr["STUDENT"].ToString() + '|';
                    data += odr["GRADE_LEVEL"].ToString() + '|';
                    data += odr["TEACHER"].ToString() + '^';

                }

                con.Close();
                con.Dispose();
            }
            catch (Exception ex)
            {
                throw;
            }
            return data;
        }

        [WebMethod]
        public static string stgrade(string gra)
        {
            string sql = string.Empty;
            string data = string.Empty;
            OracleConnection con = new OracleConnection();
            con.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings["conn"].ConnectionString;
            try
            {

                        if (con.State == ConnectionState.Closed)
                        {
                            con.Open();
                        }


                        //6 grade list
                sql = "SELECT S.STUDENT_NUMBER,S.FIRST_NAME||' '||S.LAST_NAME AS STUDENT,S.GRADE_LEVEL,T.FIRST_NAME||' '||T.LAST_NAME AS TEACHER FROM STUDENTS S";
                sql += " LEFT JOIN CC CO ON S.ID = CO.STUDENTID";
                sql += " LEFT JOIN TEACHERS T ON CO.TEACHERID = T.ID";
                sql += " LEFT JOIN COURSES C ON CO.COURSE_NUMBER = C.COURSE_NUMBER";
                if (Convert.ToInt32(gra) == 6 || Convert.ToInt32(gra) == 7 || Convert.ToInt32(gra) == 8)
                {
                    sql += " WHERE S.ENROLL_STATUS=0 AND C.COURSE_NAME LIKE '%Advisory%' AND CO.TERMID=2800 AND S.GRADE_LEVEL='" + gra + "'";
                }
              
                else if (Convert.ToInt32(gra) == 9 || Convert.ToInt32(gra) == 10 || Convert.ToInt32(gra) == 11 || Convert.ToInt32(gra) == 12)
                {
                    sql += "  WHERE S.ENROLL_STATUS = 0 AND C.COURSE_NAME LIKE '%Bohio%' AND CO.TERMID = 2800 AND S.GRADE_LEVEL = '" + gra + "'";
                }
                   else if (gra == "5" || gra== "4" || gra == "3" || gra == "2" || gra == "1" || gra == "0" || gra == "-1")
                {
                    sql += "  WHERE S.ENROLL_STATUS = 0 AND C.COURSE_NAME LIKE '%HR%' AND CO.TERMID = 2800 AND S.GRADE_LEVEL = '" + gra + "'";
                }

                sql += " ORDER BY TEACHER,STUDENT";




                OracleCommand cmd = new OracleCommand(sql, con);
                        OracleDataReader odr = cmd.ExecuteReader();
                        while (odr.Read())
                        {
                    data += odr["STUDENT_NUMBER"].ToString() + '|';
                    data += odr["STUDENT"].ToString() + '|';
                    data += odr["GRADE_LEVEL"].ToString() + '|';
                    data += odr["TEACHER"].ToString() + '^';
                    
                        }

                con.Close();
                        con.Dispose();
                    }
            catch (Exception ex)
            {
                throw ex;
            }
            return data;
        }


        [WebMethod]
        public static string ES5REPORTCARD(string stnum, string grade)
        {
            string sql = string.Empty;

            string fname = string.Empty;
            string fileName = string.Empty;
            OracleConnection con = new OracleConnection();
            con.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings["conn"].ConnectionString;
            Document documento = new Document(PageSize.LETTER, 10, 10, 5, 5);
            try
            {

                if (stnum.IndexOf(';') > -1)
                {
                    var stnumb = stnum.Split(';');
                    fname = "ES_" + grade+"Gr_ReportCard_" + DateTime.Now.DayOfYear + DateTime.Now.Month + DateTime.Now.Day + DateTime.Now.Millisecond + ".pdf";
                    fileName = HttpContext.Current.Server.MapPath("~/RepoFiles/" + fname);
                    PdfWriter.GetInstance(documento, new FileStream(fileName, FileMode.Create));
                    documento.Open();

                    for (int h = 0; h < stnumb.Length; h++)
                    {
                       
                        string T1DATA = string.Empty;
                        string T1AD = string.Empty;
                        string T1ESP = string.Empty;

                        if (con.State == ConnectionState.Closed)
                        {
                            con.Open();
                             }


                        sql = " CREATE OR REPLACE VIEW ES_VISTA";
                        sql += " AS WITH X AS(";
                        sql += " SELECT IDENTIFIER, STDID, STDDESC, stcourse, STUDENT, STUDENT_NUMBER, STDCID, STID, GRADE_LEVEL, SUBJECTAREA FROM(";
                        sql += " SELECT ST.STANDARDID STDID, ST.IDENTIFIER, TO_CHAR(TRANSIENTCOURSELIST) stcourse, ST.SUBJECTAREA, ST.NAME STDDESC FROM STANDARD ST";
                        if (grade == "2")
                        {
                            sql += " WHERE TO_CHAR(TRANSIENTCOURSELIST)  IN('" + grade + "Ma', '" + grade + "SS', '" + grade + "LA', '" + grade + "SLA', '" + grade + "HR', '" + grade + "SC', '" + grade + "PE')";
                        }
                        else if (grade == "PK")
                        {
                            sql += " WHERE TO_CHAR(TRANSIENTCOURSELIST)  IN('" + grade + "MA','" + grade + "LA', '" + grade + "SLA', '" + grade + "HR', '" + grade + "SC," + grade + "SS', '" + grade + "PE')";
                        }
                        else
                        {
                            sql += " WHERE TO_CHAR(TRANSIENTCOURSELIST)  IN('" + grade + "MA', '" + grade + "SS', '" + grade + "LA', '" + grade + "SLA', '" + grade + "HR', '" + grade + "SC', '" + grade + "PE')";
                        }
                        sql += " AND ST.YEARID = 28 AND ST.STANDARDID NOT IN(17299, 17293, 17923)AND isassignmentallowed = 1 AND ISACTIVE = 1)";
                        sql += " CROSS JOIN";
                        sql += " (";
                        sql += " SELECT FIRST_NAME || ' ' || LAST_NAME STUDENT, STUDENT_NUMBER, STUDENTS.DCID AS STDCID, ID AS STID, GRADE_LEVEL FROM STUDENTS";
                        if (grade == "K")
                        {
                            sql += " WHERE GRADE_LEVEL=0 AND ENROLL_STATUS = 0 AND STUDENT_NUMBER =" + stnumb[h] + ")";
                        }
                        else if (grade == "PK")
                        {
                            sql += " WHERE GRADE_LEVEL=-1 AND ENROLL_STATUS = 0 AND STUDENT_NUMBER =" + stnumb[h] + ")";
                        }
                        else
                        {
                            sql += " WHERE GRADE_LEVEL='" + grade + "' AND ENROLL_STATUS = 0 AND STUDENT_NUMBER =" + stnumb[h] + ")";
                        }
                        sql += " )";
                        sql += " SELECT IDENTIFIER,STDID,STDDESC,STCOURSE,STUDENT,STUDENT_NUMBER,STDCID,STID,GRADE_LEVEL,SUBJECTAREA FROM X";

                        OracleCommand cmdV1 = new OracleCommand(sql, con);
                        cmdV1.ExecuteNonQuery();


                        sql = " WITH MQUERY AS(SELECT IDENTIFIER, STDID, STDDESC, STCOURSE, STUDENT, STUDENT_NUMBER, STDCID, STID, GRADE_LEVEL,";
                        sql += " SUBJECTAREA, SG.STORECODE, SG.STANDARDGRADE, T.LASTFIRST TEACHER";
                        sql += " FROM ES_VISTA";
                        sql += " LEFT JOIN STANDARDGRADESECTION SG ON STDCID = SG.STUDENTSDCID AND STDID = SG.STANDARDID AND SG.STANDARDID IS NOT NULL AND SG.STORECODE IN('T1', 'T2', 'T3') AND SG.STANDARDGRADE<>'--'";
                        sql += " LEFT JOIN CC CO ON STID = CO.STUDENTID AND STCOURSE = CO.COURSE_NUMBER  AND CO.ORIGSECTIONID = 0";
                        sql += " LEFT JOIN TEACHERS T ON CO.TEACHERID = T.ID)";
                        sql += " SELECT DISTINCT IDENTIFIER,STID,STUDENT_NUMBER,STUDENT,GRADE_LEVEL,GRADE_LEVEL,TEACHER,STCOURSE,SUBJECTAREA,STDDESC";
                        sql += " ,(SELECT DISTINCT Y.STANDARDGRADE FROM MQUERY y WHERE y.IDENTIFIER = M.IDENTIFIER AND Y.STORECODE = 'T1') T1";
                        sql += " ,(SELECT DISTINCT Y.STANDARDGRADE FROM MQUERY y WHERE y.IDENTIFIER = M.IDENTIFIER AND Y.STORECODE = 'T2') T2";
                        sql += " ,(SELECT DISTINCT Y.STANDARDGRADE FROM MQUERY y WHERE y.IDENTIFIER = M.IDENTIFIER AND Y.STORECODE = 'T3') T3";
                        sql += " FROM MQUERY M";

                        sql += " ORDER BY CASE";
                        sql += " WHEN STCOURSE LIKE '%" + grade + "HR%' THEN 1";
                        sql += " WHEN STCOURSE LIKE '%" + grade + "LA%' THEN 2";
                        if (grade == "2")
                        {
                            sql += " WHEN STCOURSE LIKE '%" + grade + "Ma%' THEN 3";
                        }
                        else
                        {
                            sql += " WHEN STCOURSE LIKE '%" + grade + "MA%' THEN 3";
                        }
                        sql += " WHEN STCOURSE LIKE '%" + grade + "SLA%' THEN 4";
                        sql += " WHEN STCOURSE LIKE '%" + grade + "SS%' THEN 5";
                        sql += " WHEN STCOURSE LIKE '%" + grade + "SC%' THEN 6";
                        sql += " WHEN STCOURSE LIKE '%" + grade + "PE%' THEN 7";
                        sql += " END,IDENTIFIER ASC";




                        //sql = " WITH X AS(";
                        //sql += " SELECT IDENTIFIER, STDID, STDDESC, stcourse, STUDENT, STUDENT_NUMBER, STDCID, STID, GRADE_LEVEL, SUBJECTAREA FROM(";
                        //sql += " SELECT ST.STANDARDID STDID, ST.IDENTIFIER, TO_CHAR(TRANSIENTCOURSELIST) stcourse, ST.SUBJECTAREA, ST.NAME STDDESC FROM STANDARD ST";
                        //if (grade == "2")
                        //{
                        //    sql += " WHERE TO_CHAR(TRANSIENTCOURSELIST)  IN('" + grade + "Ma', '" + grade + "SS', '" + grade + "LA', '" + grade + "SLA', '" + grade + "HR', '" + grade + "SC', '" + grade + "PE')";
                        //}
                        //else
                        //{
                        //    sql += " WHERE TO_CHAR(TRANSIENTCOURSELIST)  IN('" + grade + "MA', '" + grade + "SS', '" + grade + "LA', '" + grade + "SLA', '" + grade + "HR', '" + grade + "SC', '" + grade + "PE')";
                        //}
                        //sql += " AND ST.YEARID = 28 AND ST.STANDARDID NOT IN(17299, 17293, 17923)AND isassignmentallowed = 1 AND ISACTIVE = 1)";

                        //sql += " CROSS JOIN";
                        //sql += " (";
                        //sql += " SELECT FIRST_NAME || ' ' || LAST_NAME STUDENT, STUDENT_NUMBER, STUDENTS.DCID AS STDCID, ID AS STID, GRADE_LEVEL FROM STUDENTS";
                        //if (grade == "K")
                        //{
                        //    sql += " WHERE GRADE_LEVEL=0 AND ENROLL_STATUS = 0 AND STUDENT_NUMBER =" + stnumb[h] + ")";
                        //}
                        //else if (grade == "PK")
                        //{
                        //    sql += " WHERE GRADE_LEVEL=-1 AND ENROLL_STATUS = 0 AND STUDENT_NUMBER =" + stnumb[h] + ")";
                        //}
                        //else
                        //{
                        //    sql += " WHERE GRADE_LEVEL='" + grade + "' AND ENROLL_STATUS = 0 AND STUDENT_NUMBER =" + stnumb[h] + ")";
                        //}
                        //sql += " )";
                        //sql += " SELECT STID, STUDENT_NUMBER, STUDENT, T.LASTFIRST TEACHER, IDENTIFIER, stcourse,X.GRADE_LEVEL,X.SUBJECTAREA, STDDESC";
                        //sql += " ,(SELECT  SG.STANDARDGRADE FROM X y WHERE y.IDENTIFIER = X.IDENTIFIER AND SG.STORECODE = 'T1') T1";
                        //sql += " ,(SELECT  SG.STANDARDGRADE FROM X y WHERE y.IDENTIFIER = X.IDENTIFIER AND SG.STORECODE = 'T2') T2";
                        //sql += " ,(SELECT  SG.STANDARDGRADE FROM X y WHERE y.IDENTIFIER = X.IDENTIFIER AND SG.STORECODE = 'T3') T3";

                        //sql += " FROM X";
                        //sql += " LEFT JOIN STANDARDGRADESECTION SG ON STDCID = SG.STUDENTSDCID  AND STDID = SG.STANDARDID";
                        //sql += " AND SG.STANDARDID IS NOT NULL AND SG.STORECODE IN ('T1') AND SG.YEARID = 28 AND SG.STANDARDGRADE <> '--'";
                        //sql += " LEFT JOIN CC CO ON STID = CO.STUDENTID AND STCOURSE = CO.COURSE_NUMBER  AND CO.ORIGSECTIONID = 0";
                        //sql += " LEFT JOIN TEACHERS T ON CO.TEACHERID = T.ID";
                        //sql += " ORDER BY CASE";
                        //sql += " WHEN X.STCOURSE LIKE '%" + grade + "HR%' THEN 1";
                        //sql += " WHEN X.STCOURSE LIKE '%" + grade + "LA%' THEN 2";
                        //if (grade == "2")
                        //{
                        //    sql += " WHEN X.STCOURSE LIKE '%" + grade + "Ma%' THEN 3";
                        //}
                        //else
                        //{
                        //    sql += " WHEN X.STCOURSE LIKE '%" + grade + "MA%' THEN 3";
                        //}
                        //sql += " WHEN X.STCOURSE LIKE '%" + grade + "SLA%' THEN 4";
                        //sql += " WHEN X.STCOURSE LIKE '%" + grade + "SS%' THEN 5";
                        //sql += " WHEN X.STCOURSE LIKE '%" + grade + "SC%' THEN 6";
                        //sql += " WHEN X.STCOURSE LIKE '%" + grade + "PE%' THEN 7";
                        //sql += " END,IDENTIFIER ASC";





                        OracleCommand cmd1 = new OracleCommand(sql, con);
                        OracleDataReader odr1 = cmd1.ExecuteReader();
                        while (odr1.Read())
                        {
                            T1DATA += odr1["STUDENT_NUMBER"].ToString() + '|';
                            T1DATA += odr1["STUDENT"].ToString() + '|';
                            T1DATA += odr1["GRADE_LEVEL"].ToString() + '|';
                            T1DATA += odr1["TEACHER"].ToString() + '|';
                            T1DATA += odr1["STCOURSE"].ToString() + '|';
                            T1DATA += odr1["SUBJECTAREA"].ToString() + '|';
                            T1DATA += odr1["STDDESC"].ToString() + '|';
                            T1DATA += odr1["T1"].ToString() + '|';
                            T1DATA += odr1["T2"].ToString() + '|';
                            T1DATA += odr1["T3"].ToString() + '|';
                            T1DATA += odr1["STID"].ToString() + '|';
                            T1DATA += odr1["IDENTIFIER"].ToString() + '^';

                        }


                        //sql += " CREATE OR REPLACE VIEW ES_VISTA_ESP";
                        //sql += " AS WITH X AS(";
                        //sql += " SELECT IDENTIFIER, STDID, STDDESC, stcourse, STUDENT, STUDENT_NUMBER, STDCID, STID, GRADE_LEVEL, SUBJECTAREA FROM(";
                        //sql += " SELECT ST.STANDARDID STDID, ST.IDENTIFIER, TO_CHAR(TRANSIENTCOURSELIST) stcourse, ST.SUBJECTAREA, ST.NAME STDDESC FROM STANDARD ST";
                        //if (grade == "2")
                        //{
                        //    sql += " WHERE TO_CHAR(TRANSIENTCOURSELIST)  IN('" + grade + "TECH', '" + grade + "Art', '" + grade + "Mus')";
                        //}
                        //else { 
                        //    sql += " WHERE TO_CHAR(TRANSIENTCOURSELIST)  IN('" + grade + "TECH', '" + grade + "ART', '" + grade + "MUS')";
                        //{
                        //sql += " AND ST.YEARID = 28 AND ST.STANDARDID NOT IN(17299, 17293, 17923)AND isassignmentallowed = 1 AND ISACTIVE = 1)";
                        //sql += " CROSS JOIN";
                        //sql += " (";
                        //sql += " SELECT FIRST_NAME || ' ' || LAST_NAME STUDENT, STUDENT_NUMBER, STUDENTS.DCID AS STDCID, ID AS STID, GRADE_LEVEL FROM STUDENTS";
                        //if (grade == "K")
                        //{
                        //    sql += " WHERE GRADE_LEVEL=0 AND ENROLL_STATUS = 0 AND STUDENT_NUMBER =" + stnumb[h] + ")";
                        //}
                        //else if (grade == "PK")
                        //{
                        //    sql += " WHERE GRADE_LEVEL=-1 AND ENROLL_STATUS = 0 AND STUDENT_NUMBER =" + stnumb[h] + ")";
                        //}
                        //else
                        //{
                        //    sql += " WHERE GRADE_LEVEL='" + grade + "' AND ENROLL_STATUS = 0 AND STUDENT_NUMBER =" + stnumb[h] + ")";
                        //}
                        //sql += " )";
                        //sql += " SELECT IDENTIFIER, STDID, STDDESC, STCOURSE, STUDENT, STUDENT_NUMBER, STDCID, STID, GRADE_LEVEL, SUBJECTAREA FROM X;";

                        //OracleCommand cmdV2 = new OracleCommand(sql, con);
                        //cmdV2.ExecuteNonQuery();

                        //sql += " SELECT TEACHER, STCOURSE, LISTAGG(T1,',') WITHIN GROUP (ORDER BY STDDESC) T1,LISTAGG(T2, ',') WITHIN GROUP (ORDER BY STDDESC)T2,LISTAGG(T3, ',') WITHIN GROUP (ORDER BY STDDESC)T3 FROM (";
                        //sql += " WITH MQUERY AS(SELECT IDENTIFIER, STDID, STDDESC, STCOURSE, STUDENT, STUDENT_NUMBER, STDCID, STID, GRADE_LEVEL,";
                        //sql += " SUBJECTAREA, SG.STORECODE, SG.STANDARDGRADE, T.LASTFIRST TEACHER";
                        //sql += " FROM ES_VISTA_ESP";
                        //sql += " LEFT JOIN STANDARDGRADESECTION SG ON STDCID = SG.STUDENTSDCID AND sg.yearid = 28 and STDID = SG.STANDARDID AND STANDARDGRADE <> '--'";
                        //sql += "  AND SG.STANDARDID IS NOT NULL AND SG.STORECODE IN('T1', 'T2', 'T3')";
                        //sql += " LEFT JOIN CC CO ON STID = CO.STUDENTID AND STCOURSE = CO.COURSE_NUMBER  AND CO.ORIGSECTIONID = 0";
                        //sql += " LEFT JOIN TEACHERS T ON CO.TEACHERID = T.ID)";
                        //sql += " SELECT DISTINCT TEACHER, STCOURSE, STDDESC";
                        //sql += " , (SELECT distinct Y.STANDARDGRADE FROM MQUERY y WHERE y.IDENTIFIER = M.IDENTIFIER AND Y.STORECODE = 'T1') T1";
                        //sql += " ,(SELECT distinct Y.STANDARDGRADE FROM MQUERY y WHERE y.IDENTIFIER = M.IDENTIFIER AND Y.STORECODE = 'T2') T2";
                        //sql += " ,(SELECT distinct Y.STANDARDGRADE FROM MQUERY y WHERE y.IDENTIFIER = M.IDENTIFIER AND Y.STORECODE = 'T3') T3";
                        //sql += " FROM MQUERY M";
                        //sql += " )";
                        //sql += " GROUP BY TEACHER,STCOURSE";
                        //sql += " ORDER BY CASE";
                        //sql += " WHEN STCOURSE LIKE '%" + grade + "TECH%' THEN 1";
                        //if (grade == "2")
                        //{
                        //    sql += " WHEN STCOURSE LIKE '%" + grade + "Art%' THEN 2";
                        //    sql += " WHEN STCOURSE LIKE '%" + grade + "Mus%' THEN 3";
                        //}
                        //else
                        //{
                        //    sql += " WHEN STCOURSE LIKE '%" + grade + "ART%' THEN 2";
                        //    sql += " WHEN STCOURSE LIKE '%" + grade + "MUS%' THEN 3";
                        //}
                        //    sql += " END";

                        sql = "WITH X AS(";
                        sql += " SELECT IDENTIFIER, STANDARDID, STDDESC, stcourse, LASTFIRST, STUDENT_NUMBER, STDCID, STID, GRADE_LEVEL, SUBJECTAREA FROM(";
                        sql += " SELECT ST.STANDARDID, ST.IDENTIFIER, TO_CHAR(TRANSIENTCOURSELIST) stcourse, ST.SUBJECTAREA, ST.NAME STDDESC FROM STANDARD ST";
                        if (grade == "2")
                        {
                            sql += " WHERE TO_CHAR(TRANSIENTCOURSELIST)  IN('" + grade + "TECH', '" + grade + "Art', '" + grade + "Mus')";
                        }else
                        {
                            sql += " WHERE TO_CHAR(TRANSIENTCOURSELIST)  IN('" + grade + "TECH', '" + grade + "ART', '" + grade + "MUS')";
                        }
                        sql += " AND ST.YEARID = 28 AND STANDARDID NOT IN(17299, 17293)AND isassignmentallowed = 1)";
                        sql += " CROSS JOIN";
                        sql += " (";
                        sql += " SELECT LASTFIRST, STUDENT_NUMBER, STUDENTS.DCID AS STDCID, ID AS STID, GRADE_LEVEL FROM STUDENTS";
                        if (grade == "K")
                        {
                            sql += " WHERE GRADE_LEVEL=0 AND ENROLL_STATUS = 0 AND STUDENT_NUMBER =" + stnumb[h] + ")";
                        }
                        else if (grade == "PK")
                        {
                            sql += " WHERE GRADE_LEVEL=-1 AND ENROLL_STATUS = 0 AND STUDENT_NUMBER =" + stnumb[h] + ")";
                        }
                        else
                        {
                            sql += " WHERE GRADE_LEVEL='" + grade + "' AND ENROLL_STATUS = 0 AND STUDENT_NUMBER =" + stnumb[h] + ")";
                        }
                        sql += " )";
                        sql += " SELECT STUDENT, TEACHER, STCOURSE,(CASE WHEN STORECODE = 'T1' THEN T_GRADE || '/' || U_GRADE ELSE NULL END)T1";
                        sql += " ,(CASE WHEN STORECODE = 'T2' THEN T_GRADE|| '/' || U_GRADE ELSE NULL END)T2";
                        sql += " ,(CASE WHEN STORECODE = 'T3' THEN T_GRADE|| '/' || U_GRADE ELSE NULL END)T3";
                        sql += " FROM(SELECT X.STDDESC, X.STID, X.STUDENT_NUMBER, X.LASTFIRST STUDENT, T.LASTFIRST TEACHER, X.STCOURSE, SG.STANDARDGRADE, SG.STORECODE";
                        sql += " FROM X";
                        sql += " LEFT JOIN STANDARDGRADESECTION SG ON X.STDCID = SG.STUDENTSDCID AND X.STANDARDID = SG.STANDARDID AND SG.STANDARDID IS NOT NULL AND SG.STORECODE IN('T1') AND SG.STANDARDGRADE<>'--'";
                        sql += " LEFT JOIN CC CO ON X.STID = CO.STUDENTID AND X.STCOURSE = CO.COURSE_NUMBER  AND CO.ORIGSECTIONID = 0";
                        sql += " LEFT JOIN TEACHERS T ON CO.TEACHERID = T.ID";
                        sql += " )";
                        sql += " PIVOT(MAX(standardgrade) AS grade FOR(STDDESC) IN('Understands concepts and uses skills.' AS U,'Follows Tribes® Agreements.' AS T))";
                        sql += " ORDER BY STUDENT,STCOURSE DESC";


                        OracleCommand cmd2 = new OracleCommand(sql, con);
                        OracleDataReader odr2 = cmd2.ExecuteReader();
                        while (odr2.Read())
                        {
                            T1ESP += odr2["STCOURSE"].ToString() + '|';
                            T1ESP += odr2["TEACHER"].ToString() + '|';
                            T1ESP += odr2["T1"].ToString() + '|';
                            T1ESP += odr2["T2"].ToString() + '|';
                            T1ESP += odr2["T3"].ToString() + '^';
                        }


                        sql = " SELECT LISTAGG(T1,',') WITHIN GROUP (ORDER BY STUDENT) T1,LISTAGG(T2, ',') WITHIN GROUP (ORDER BY STUDENT) T2,LISTAGG(T3, ',') WITHIN GROUP (ORDER BY STUDENT) T3 FROM (SELECT STUDENT,";
                        sql += " (CASE WHEN ABBRE = 'T1' THEN COUNT(ABSENCE) || ',' || COUNT(TARDI) END)T1,(CASE WHEN ABBRE = 'T2' THEN COUNT(ABSENCE) || ',' || COUNT(TARDI) END)T2";
                        sql += " ,(CASE WHEN ABBRE = 'T3' THEN COUNT(ABSENCE) || ',' || COUNT(TARDI) END)T3 FROM (SELECT DISTINCT S.LASTFIRST STUDENT,";
                        sql += " AC.ATT_CODE, (CASE WHEN(AC.ATT_CODE = 'EA' OR AC.ATT_CODE = 'UA') THEN AC.PRESENCE_STATUS_CD END) ABSENCE,";
                        sql += " (CASE  WHEN(AC.ATT_CODE = 'ET' OR AC.ATT_CODE = 'UT') THEN AC.PRESENCE_STATUS_CD END) TARDI, AT.ATT_DATE,T.ABBREVIATION ABBRE FROM ATTENDANCE AT";
                        sql += " LEFT JOIN STUDENTS S ON AT.STUDENTID = S.ID";
                        sql += " LEFT JOIN ATTENDANCE_CODE AC ON AT.ATTENDANCE_CODEID = AC.ID";
                        sql += " LEFT JOIN TERMS T ON AT.YEARID = T.YEARID";
                        sql += " WHERE AT.YEARID = 28 AND AT.ATT_DATE BETWEEN (T.FIRSTDAY)AND(T.LASTDAY) AND T.ABBREVIATION IN('T1') AND S.STUDENT_NUMBER = " + stnumb[h] + "";
                        sql += " )";
                        sql += " GROUP BY STUDENT,ABBRE)";
                        sql += " GROUP BY STUDENT";


                        //sql = " SELECT STUDENT, COUNT(ABSENCE) ABSE, COUNT(TARDI) TARD FROM(SELECT DISTINCT S.LASTFIRST STUDENT,";
                        //sql += " AC.ATT_CODE, (CASE WHEN(AC.ATT_CODE = 'EA' OR AC.ATT_CODE = 'UA') THEN AC.PRESENCE_STATUS_CD END) ABSENCE,";
                        //sql += " (CASE  WHEN(AC.ATT_CODE = 'ET' OR AC.ATT_CODE = 'UT' ) THEN AC.PRESENCE_STATUS_CD END) TARDI, AT.ATT_DATE FROM ATTENDANCE AT";
                        //sql += " LEFT JOIN STUDENTS S ON AT.STUDENTID = S.ID";
                        //sql += " LEFT JOIN ATTENDANCE_CODE AC ON AT.ATTENDANCE_CODEID = AC.ID";
                        //sql += " WHERE AT.YEARID = 28 AND AT.ATT_DATE <= CURRENT_DATE AND S.STUDENT_NUMBER = " + stnumb[h] + "";
                        //sql += " )";
                        //sql += " GROUP BY STUDENT";


                        string ABTAR = string.Empty;
                        OracleCommand cmd3 = new OracleCommand(sql, con);
                        OracleDataReader odr3 = cmd3.ExecuteReader();
                        while (odr3.Read())
                        {
                            ABTAR += odr3["T1"].ToString() + '|';
                            ABTAR += odr3["T2"].ToString() + '|';
                            ABTAR += odr3["T3"].ToString() + '|';
                        }


                        sql = " SELECT S.LASTFIRST STUDENT, ST.COMMENTVALUE FROM STANDARDGRADESECTIONCOMMENT ST";
                        sql += " LEFT JOIN STUDENTS S ON ST.STUDENTSDCID = S.DCID";
                        sql += " LEFT JOIN STANDARDGRADESECTION SG ON ST.STANDARDGRADESECTIONID = SG.STANDARDGRADESECTIONID";
                        if (grade == "K") { 
                        sql += " WHERE ST.YEARID = 28 AND S.GRADE_LEVEL='0' AND SG.STORECODE='T1' AND S.STUDENT_NUMBER = " + stnumb[h] + "";
                        }else if (grade == "PK")
                        {
                            sql += " WHERE ST.YEARID = 28 AND S.GRADE_LEVEL='-1' AND SG.STORECODE='T1' AND S.STUDENT_NUMBER = " + stnumb[h] + "";
                        }
                        else
                        {
                            sql += " WHERE ST.YEARID = 28 AND S.GRADE_LEVEL='" + grade + "' AND SG.STORECODE='T1' AND S.STUDENT_NUMBER = " + stnumb[h] + "";
                        }

                        string comm = string.Empty;
                        OracleCommand cmd4 = new OracleCommand(sql, con);
                        OracleDataReader odr4 = cmd4.ExecuteReader();
                        while (odr4.Read())
                        {
                            comm = odr4["COMMENTVALUE"].ToString();
                        }

                        if (T1DATA != "")
                        {


                            var stTable = T1DATA.Split('^');
                            var HT = "";
                            var std = "";
                            var stn = "";
                            var stgd = "";
                            var stid = "";
                            for (int i = 0; i < stTable.Length; i++)
                            {
                                var hr = stTable[i].Split('|');
                                if (hr[11].Split('.')[2] == "HWr")
                                {
                                    HT = hr[3];
                                    std = hr[0];
                                    stn = hr[1];
                                    stgd = hr[2];
                                    stid = hr[10];
                                    break;
                                }
                                else if (hr[4] == "PKHR")
                                {
                                    HT = hr[3];
                                    std = hr[0];
                                    stn = hr[1];
                                    stgd = hr[2];
                                    stid = hr[10];
                                    break;
                                }

                            }



                            iTextSharp.text.Image Imagen = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("~/img/WLOGO.jpg"));
                            Imagen.ScalePercent(2.5f);

                            iTextSharp.text.Image foto;
                          //  fileName("file://cms03pws/e$/program%20files/powerschool/data/picture/student/" + stid.Substring(stid.Length - Math.Min(2, stid.Length)) + "/" + stid + "/ph.jpeg");
                           // {
                                foto = iTextSharp.text.Image.GetInstance("file://cms03pws/e$/program%20files/powerschool/data/picture/student/" + stid.Substring(stid.Length - Math.Min(2, stid.Length)) + "/" + stid + "/ph.jpeg");
                                foto.ScalePercent(27f);
                           // }else
                            //{
                            //    foto = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("~/img/perfil.png"));
                            //    foto.ScalePercent(27f);
                            //}


                            PdfPTable HeadT = new PdfPTable(16);
                            HeadT.HorizontalAlignment = Element.ALIGN_CENTER;
                            HeadT.WidthPercentage = 100;

                            PdfPCell logo = new PdfPCell(Imagen);
                            logo.Colspan = 9;
                            logo.Border = 0;
                            logo.HorizontalAlignment = Element.ALIGN_LEFT;
                            logo.Rowspan = 3;
                            logo.Padding = 3;
                            HeadT.AddCell(logo);


                            PdfPCell HS = new PdfPCell(new Phrase("ELEMENTARY SCHOOL Report Card", new Font(Font.FontFamily.HELVETICA, 14, Font.BOLD, new BaseColor(135, 0, 27))));
                            HS.HorizontalAlignment = Element.ALIGN_BOTTOM;
                            HS.Colspan = 7;
                            HS.Border = 0;
                            HeadT.AddCell(HS);

                            PdfPCell SQ1 = new PdfPCell(new Phrase("School Year 2018-19 Trimester 1", new Font(Font.FontFamily.HELVETICA, 14, Font.BOLD, BaseColor.BLACK)));
                            SQ1.HorizontalAlignment = Element.ALIGN_BOTTOM;
                            SQ1.Colspan = 7;
                            SQ1.Border = 0;
                            HeadT.AddCell(SQ1);

                            PdfPCell Pub = new PdfPCell(new Phrase("Published " + DateTime.Now.ToString("MMMM dd, yyyy"), new Font(Font.FontFamily.HELVETICA, 12, Font.ITALIC, BaseColor.BLACK)));
                            Pub.Colspan = 7;
                            Pub.HorizontalAlignment = Element.ALIGN_BOTTOM;
                            Pub.Rowspan = 2;
                            Pub.Border = 0;
                            HeadT.AddCell(Pub);

                            PdfPCell bar1 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                            bar1.HorizontalAlignment = Element.ALIGN_LEFT;
                            bar1.Border = 0;
                            bar1.Colspan = 16;
                            bar1.BackgroundColor = new BaseColor(135, 0, 27);
                            HeadT.AddCell(bar1);

                            PdfPCell stinfo = new PdfPCell(new Phrase("Student Name: " + stn, new Font(Font.FontFamily.HELVETICA, 11, Font.BOLD, BaseColor.BLACK)));
                            stinfo.HorizontalAlignment = Element.ALIGN_LEFT;
                            stinfo.Border = 0;
                            stinfo.Colspan = 7;
                            stinfo.PaddingTop = 5;
                            HeadT.AddCell(stinfo);

                            PdfPCell stfoto = new PdfPCell(foto);
                            stfoto.HorizontalAlignment = Element.ALIGN_LEFT;
                            stfoto.Colspan = 2;
                            stfoto.Border = 0;
                            stfoto.Rowspan = 3;
                            stfoto.PaddingTop = 0.5f;
                            stfoto.PaddingBottom = 1f;
                            HeadT.AddCell(stfoto);

                            PdfPCell messag = new PdfPCell(new Phrase("The purpose of this report is to communicate student achievement in relationship" + Environment.NewLine + "to trimester goals as well as what is required for future progress toward them.", new Font(Font.FontFamily.HELVETICA, 11, Font.NORMAL, BaseColor.BLACK)));
                            messag.HorizontalAlignment = Element.ALIGN_LEFT;
                            messag.Border = 0;
                            messag.Colspan = 7;
                            messag.PaddingTop = 2;
                            messag.PaddingBottom = 5;
                            messag.Rowspan = 3;
                            HeadT.AddCell(messag);

                            PdfPCell grad = new PdfPCell(new Phrase("Grade: " + grade, new Font(Font.FontFamily.HELVETICA, 11, Font.BOLD, BaseColor.BLACK)));
                            grad.HorizontalAlignment = Element.ALIGN_LEFT;
                            grad.Border = 0;
                            grad.Colspan = 5;
                            HeadT.AddCell(grad);

                            PdfPCell stinu = new PdfPCell(new Phrase("StudentID: " + std, new Font(Font.FontFamily.HELVETICA, 5, Font.BOLD, BaseColor.WHITE)));
                            stinu.HorizontalAlignment = Element.ALIGN_LEFT;
                            stinu.Border = 0;
                            stinu.Colspan = 5;
                            stinu.PaddingBottom = 3;
                            HeadT.AddCell(stinu);

                            PdfPCell HR = new PdfPCell(new Phrase("Homeroom: " + HT, new Font(Font.FontFamily.HELVETICA, 11, Font.BOLD, BaseColor.BLACK)));
                            HR.HorizontalAlignment = Element.ALIGN_LEFT;
                            HR.Border = 0;
                            HR.Colspan = 10;
                            HR.PaddingBottom = 5;
                            HeadT.AddCell(HR);


                            PdfPCell bar2 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                            bar2.HorizontalAlignment = Element.ALIGN_LEFT;
                            bar2.Border = 0;
                            bar2.Colspan = 16;
                            bar2.BackgroundColor = new BaseColor(135, 0, 27);
                            HeadT.AddCell(bar2);

                            //Legend
                            PdfPTable legendTable = new PdfPTable(18);
                            legendTable.HorizontalAlignment = Element.ALIGN_CENTER;
                            legendTable.WidthPercentage = 100;

                            PdfPCell cel1 = new PdfPCell(new Phrase("STANDARDS PROFICIENCY KEY ", new Font(Font.FontFamily.HELVETICA, 9F, Font.BOLD, BaseColor.BLACK)));
                            cel1.HorizontalAlignment = Element.ALIGN_LEFT;
                            cel1.Colspan = 18;
                            cel1.BorderColor = BaseColor.LIGHT_GRAY;
                            cel1.PaddingTop = 5;
                            cel1.BorderWidthTop = 0;
                            cel1.BorderWidthBottom = 1;
                            cel1.BorderWidthRight = 0;
                            cel1.BorderWidthLeft = 0;
                            legendTable.AddCell(cel1);
                            PdfPCell cel2 = new PdfPCell(new Phrase("Code", new Font(Font.FontFamily.HELVETICA, 9F, Font.BOLD, BaseColor.BLACK)));
                            cel2.HorizontalAlignment = Element.ALIGN_LEFT;
                            cel2.BorderWidthLeft = 1;
                            cel2.BorderWidthBottom = 1;
                            cel2.BorderColor = BaseColor.LIGHT_GRAY;
                            cel2.BorderWidthRight = 0;
                            cel2.BorderWidthTop = 0;
                            legendTable.AddCell(cel2);
                            PdfPCell cel3 = new PdfPCell(new Phrase("Achievement Descriptors", new Font(Font.FontFamily.HELVETICA, 9F, Font.BOLD, BaseColor.BLACK)));
                            cel3.HorizontalAlignment = Element.ALIGN_LEFT;
                            cel3.BorderWidthBottom = 1;
                            cel3.BorderWidthRight = 0;
                            cel3.BorderWidthLeft = 0;
                            cel3.BorderWidthTop = 0;
                            cel3.BorderColor = BaseColor.LIGHT_GRAY;
                            cel3.Colspan = 5;
                            legendTable.AddCell(cel3);
                            PdfPCell cel4 = new PdfPCell(new Phrase("Behavioral Descriptors", new Font(Font.FontFamily.HELVETICA, 9F, Font.BOLD, BaseColor.BLACK)));
                            cel4.HorizontalAlignment = Element.ALIGN_LEFT;
                            cel4.Colspan = 12;
                            cel4.BorderWidthLeft = 0;
                            cel4.BorderWidthRight = 1;
                            cel4.BorderWidthTop = 0;
                            cel4.BorderColor = BaseColor.LIGHT_GRAY;
                            legendTable.AddCell(cel4);
                            PdfPCell cel1d = new PdfPCell(new Phrase("5" + Environment.NewLine + "4" + Environment.NewLine + "3" + Environment.NewLine + "2" + Environment.NewLine + "1" + Environment.NewLine + "--" + Environment.NewLine + "*", new Font(Font.FontFamily.HELVETICA, 8.0F, Font.NORMAL, BaseColor.BLACK)));
                            cel1d.HorizontalAlignment = Element.ALIGN_CENTER;
                            cel1d.BorderWidthLeft = 1;
                            cel1d.BorderWidthBottom = 1;
                            cel1d.BorderWidthRight = 0;
                            cel1d.BorderWidthTop = 0;
                            cel1d.BorderColor = BaseColor.LIGHT_GRAY;
                            legendTable.AddCell(cel1d);
                            PdfPCell cel2d = new PdfPCell(new Phrase("Meets Trimester Standard with Distinction" + Environment.NewLine + "Meets Trimester Standard" + Environment.NewLine + "Nearly Meets Trimester Standard" + Environment.NewLine + "Below Trimester Standard" + Environment.NewLine + "Far Below Trimester Standard" + Environment.NewLine + "Not Assessed This Trimester" + Environment.NewLine + "Based on Modified Expectations", new Font(Font.FontFamily.HELVETICA, 8F, Font.NORMAL, BaseColor.BLACK)));
                            cel2d.HorizontalAlignment = Element.ALIGN_LEFT;
                            cel2d.Colspan = 5;
                            cel2d.BorderWidthBottom = 1;
                            cel2d.BorderWidthTop = 0;
                            cel2d.BorderWidthLeft = 0;
                            cel2d.BorderWidthRight = 0;
                            cel2d.BorderColor = BaseColor.LIGHT_GRAY;
                            legendTable.AddCell(cel2d);
                            PdfPCell cel3d = new PdfPCell(new Phrase("The student takes understandings and learning beyong trimester benchmark consistantly." + Environment.NewLine + "The student knows and/or is able to do trimester benchmark consistently" + Environment.NewLine + "The student knows and/or is able to do trimester benchmark, but not consistently" + Environment.NewLine + "The student does not know and/or unable to do trimester benchmark, but shows beginning understandings." + Environment.NewLine + "The student does not know and/or is unable to do trimester benchmark." + Environment.NewLine + "The student was not assessed on this benchmark this trimester." + Environment.NewLine + "The student was assessed based on his/her individualized educational goals.", new Font(Font.FontFamily.HELVETICA, 8F, Font.NORMAL, BaseColor.BLACK)));
                            cel3d.HorizontalAlignment = Element.ALIGN_LEFT;
                            cel3d.Colspan = 12;
                            cel3d.BorderWidthLeft = 0;
                            cel3d.BorderWidthBottom = 1;
                            cel3d.BorderWidthRight = 1;
                            cel3d.BorderColor = BaseColor.LIGHT_GRAY;
                            legendTable.AddCell(cel3d);

                            /// GRADE DETAILS
                            PdfPTable GradeTable = new PdfPTable(20);
                            GradeTable.HorizontalAlignment = Element.ALIGN_CENTER;
                            GradeTable.WidthPercentage = 100;

                            var subt = "";
                            var TAPB = "";
                            var hed = "";
                            var hed2 = "";
                            var hrw = "";
                            for (int i = 0; i < stTable.Length - 1; i++)
                            {
                                var hr = stTable[i].Split('|');
                                var idt = hr[11].Split('.')[2];
                                hed = hr[4];
                                if (hr[5] == "Handwriting")
                                {
                                    hrw = hr[5] + "|" + hr[6] + "|" + hr[7] + "|" + hr[8] + "|" + hr[9];
                                }
                                if (hed == "" + grade + "LA")
                                {
                                    TAPB = "ENGLISH LANGUAGE ARTS";
                                }
                                else if (hed == "" + grade + "MA" || hed == "" + grade + "Ma")
                                {
                                    TAPB = "MATHEMATICS";
                                }
                                else if (hed == "" + grade + "SLA")
                                {
                                    TAPB = "SPANISH LANGUAGE ARTS";
                                }
                                else if (hed == "" + grade + "HR")
                                {
                                    TAPB = "CONDUCT";
                                }

                                else if (hed == "" + grade + "SC")
                                {
                                    TAPB = "SCIENCE";
                                }
                                else if (hed == "" + grade + "SS")
                                {
                                    TAPB = "SOCIAL STUDIES";
                                }
                                else if (hed == "" + grade + "PE")
                                {
                                    TAPB = "PHYSICAL EDUCATION / HEALTH";
                                }
                                else if (hed == "PKSC,PKSS")
                                {
                                    TAPB = "SOCIAL STUDIES";
                                    hr[3] = HT;
                                }



                                if (hed != hed2)
                                {

                                    PdfPCell spa2 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 5.0F, Font.NORMAL, BaseColor.BLACK)));
                                    spa2.Border = 0;
                                    spa2.Colspan = 20;
                                    GradeTable.AddCell(spa2);

                                    PdfPCell Course = new PdfPCell(new Phrase(TAPB, new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.WHITE)));
                                    Course.HorizontalAlignment = Element.ALIGN_CENTER;
                                    Course.BorderWidth = 1F;
                                    Course.BackgroundColor = new BaseColor(135, 0, 27);
                                    Course.Colspan = 6;
                                    Course.BorderColor = BaseColor.GRAY;
                                    GradeTable.AddCell(Course);
                                    PdfPCell Teacher = new PdfPCell(new Phrase(hr[3], new Font(Font.FontFamily.HELVETICA, 11.0F, Font.ITALIC, BaseColor.BLACK)));
                                    Teacher.HorizontalAlignment = Element.ALIGN_LEFT;
                                    Teacher.BorderWidth = 1F;
                                    Teacher.Colspan = 11;
                                    Teacher.BorderColor = BaseColor.GRAY;
                                    Teacher.Border = 0;
                                    GradeTable.AddCell(Teacher);


                                    if (hr[4] == "" + grade + "HR")
                                    {

                                        PdfPCell T1 = new PdfPCell(new Phrase("T1", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                                        T1.HorizontalAlignment = Element.ALIGN_LEFT;
                                        T1.BackgroundColor = new BaseColor(135, 0, 27);
                                        T1.BorderWidth = 1F;
                                        GradeTable.AddCell(T1);
                                        PdfPCell T2 = new PdfPCell(new Phrase("T2", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                                        T2.HorizontalAlignment = Element.ALIGN_LEFT;
                                        T2.BackgroundColor = new BaseColor(135, 0, 27);
                                        T2.BorderWidth = 1F;
                                        GradeTable.AddCell(T2);
                                        PdfPCell T3 = new PdfPCell(new Phrase("T3", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                                        T3.HorizontalAlignment = Element.ALIGN_LEFT;
                                        T3.BackgroundColor = new BaseColor(135, 0, 27);
                                        T3.BorderWidth = 1F;
                                        GradeTable.AddCell(T3);
                                    }
                                    else
                                    {
                                        PdfPCell spa = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        spa.HorizontalAlignment = Element.ALIGN_LEFT;
                                        spa.BorderWidth = 1F;
                                        spa.Colspan = 3;
                                        spa.Border = 0;
                                        GradeTable.AddCell(spa);
                                    }

                                    if (idt == "LA" && hrw != "")
                                    {
                                        PdfPCell subj1 = new PdfPCell(new Phrase(hrw.Split('|')[0], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        subj1.HorizontalAlignment = Element.ALIGN_LEFT;
                                        subj1.BorderWidth = 1F;
                                        subj1.Colspan = 4;
                                        subj1.BorderColor = BaseColor.GRAY;
                                        GradeTable.AddCell(subj1);

                                        PdfPCell stnam1 = new PdfPCell(new Phrase(hrw.Split('|')[1], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        stnam1.HorizontalAlignment = Element.ALIGN_LEFT;
                                        stnam1.BorderWidth = 1F;
                                        stnam1.Colspan = 13;
                                        stnam1.BorderColor = BaseColor.GRAY;
                                        GradeTable.AddCell(stnam1);


                                        PdfPCell vt12 = new PdfPCell(new Phrase(hrw.Split('|')[2], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        vt12.HorizontalAlignment = Element.ALIGN_CENTER;
                                        vt12.BorderWidth = 1F;
                                        vt12.BorderColor = BaseColor.GRAY;
                                        GradeTable.AddCell(vt12);

                                        PdfPCell vt22 = new PdfPCell(new Phrase(hrw.Split('|')[3], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        vt22.HorizontalAlignment = Element.ALIGN_CENTER;
                                        vt22.BorderWidth = 1F;
                                        vt22.BorderColor = BaseColor.GRAY;
                                        GradeTable.AddCell(vt22);


                                        PdfPCell vt31 = new PdfPCell(new Phrase(hrw.Split('|')[4], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        vt31.HorizontalAlignment = Element.ALIGN_CENTER;
                                        vt31.BorderWidth = 1F;
                                        vt31.BorderColor = BaseColor.GRAY;
                                        GradeTable.AddCell(vt31);

                                        hrw = "";
                                    }

                                    hed2 = hr[4];

                                }

                                if (hr[5] == "Tribes" || hr[5] == "TRIBES" || hr[5] == "TribesTLCÆ")
                                {
                                    hr[5] = "Tribes® Agreements";
                                }
                                else if (hr[5] == "Listening and Speaking.")
                                {
                                    hr[5] = "Listening and Speaking";
                                }
                                else if (hr[5] == "SOCIAL/WORK DEVELOPMENT (TRIBES®)" || hr[5] == "Social/Work Development (TRIBES)")
                                {
                                    hr[5] = "Tribes® Agreements";
                                }

                                if (hr[5] != "Handwriting")
                                {

                                    if (subt != hr[5])
                                    {

                                        PdfPCell subj = new PdfPCell(new Phrase(hr[5], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        subj.HorizontalAlignment = Element.ALIGN_LEFT;
                                        subj.BorderWidth = 1F;
                                        subj.Colspan = 4;
                                        subj.BorderColor = BaseColor.GRAY;
                                        GradeTable.AddCell(subj);
                                        subt = hr[5];
                                    }
                                    else
                                    {

                                        PdfPCell subj = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        subj.HorizontalAlignment = Element.ALIGN_LEFT;
                                        subj.Border = 0;
                                        subj.Colspan = 4;
                                        GradeTable.AddCell(subj);
                                    }

                                    PdfPCell stnam = new PdfPCell(new Phrase(hr[6], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                    stnam.HorizontalAlignment = Element.ALIGN_LEFT;
                                    stnam.BorderWidth = 1F;
                                    stnam.Colspan = 13;
                                    stnam.BorderColor = BaseColor.GRAY;
                                    GradeTable.AddCell(stnam);

                                    PdfPCell vt1 = new PdfPCell(new Phrase(hr[7], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                    vt1.HorizontalAlignment = Element.ALIGN_CENTER;
                                    vt1.BorderWidth = 1F;
                                    vt1.BorderColor = BaseColor.GRAY;
                                    GradeTable.AddCell(vt1);



                                    PdfPCell vt2 = new PdfPCell(new Phrase(hr[8], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                    vt2.HorizontalAlignment = Element.ALIGN_CENTER;
                                    vt2.BorderWidth = 1F;
                                    vt2.BorderColor = BaseColor.GRAY;
                                    GradeTable.AddCell(vt2);


                                    PdfPCell vt3 = new PdfPCell(new Phrase(hr[9], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                    vt3.HorizontalAlignment = Element.ALIGN_CENTER;
                                    vt3.BorderWidth = 1F;
                                    vt3.BorderColor = BaseColor.GRAY;
                                    GradeTable.AddCell(vt3);


                                }
                            }
                            PdfPTable espTable = new PdfPTable(18);
                            espTable.HorizontalAlignment = Element.ALIGN_CENTER;
                            espTable.WidthPercentage = 100;
                            if (T1ESP != "")
                            {
                                var espT = T1ESP.Split('^');

                                PdfPCell ep1 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                                ep1.HorizontalAlignment = Element.ALIGN_LEFT;
                                ep1.Colspan = 18;
                                ep1.Border = 0;
                                espTable.AddCell(ep1);
                                PdfPCell espe = new PdfPCell(new Phrase("Fine Arts & Technology", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                                espe.HorizontalAlignment = Element.ALIGN_LEFT;
                                espe.BackgroundColor = new BaseColor(135, 0, 27);
                                espe.BorderWidth = 1F;
                                espe.Colspan = 12;
                                espTable.AddCell(espe);
                                PdfPCell ST1 = new PdfPCell(new Phrase("T1", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                                ST1.HorizontalAlignment = Element.ALIGN_CENTER;
                                ST1.BackgroundColor = new BaseColor(135, 0, 27);
                                ST1.BorderWidth = 1F;
                                ST1.Colspan = 2;
                                espTable.AddCell(ST1);
                                PdfPCell ST2 = new PdfPCell(new Phrase("T2", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                                ST2.HorizontalAlignment = Element.ALIGN_CENTER;
                                ST2.BackgroundColor = new BaseColor(135, 0, 27);
                                ST2.BorderWidth = 1F;
                                ST2.Colspan = 2;
                                espTable.AddCell(ST2);
                                PdfPCell ST3 = new PdfPCell(new Phrase("T3", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                                ST3.HorizontalAlignment = Element.ALIGN_CENTER;
                                ST3.BackgroundColor = new BaseColor(135, 0, 27);
                                ST3.BorderWidth = 1F;
                                ST3.Colspan = 2;
                                espTable.AddCell(ST3);
                                PdfPCell SP = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                                SP.HorizontalAlignment = Element.ALIGN_LEFT;
                                SP.BackgroundColor = new BaseColor(135, 0, 27);
                                SP.BorderWidth = 1F;
                                SP.Colspan = 12;
                                espTable.AddCell(SP);
                                PdfPCell U1 = new PdfPCell(new Phrase("U", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                                U1.HorizontalAlignment = Element.ALIGN_CENTER;
                                U1.BackgroundColor = new BaseColor(135, 0, 27);
                                U1.BorderWidth = 1F;
                                espTable.AddCell(U1);
                                PdfPCell SPT1 = new PdfPCell(new Phrase("T", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                                SPT1.HorizontalAlignment = Element.ALIGN_CENTER;
                                SPT1.BackgroundColor = new BaseColor(135, 0, 27);
                                SPT1.BorderWidth = 1F;
                                espTable.AddCell(SPT1);
                                PdfPCell U2 = new PdfPCell(new Phrase("U", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                                U2.HorizontalAlignment = Element.ALIGN_CENTER;
                                U2.BackgroundColor = new BaseColor(135, 0, 27);
                                U2.BorderWidth = 1F;
                                espTable.AddCell(U2);
                                PdfPCell SPT2 = new PdfPCell(new Phrase("T", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                                SPT2.HorizontalAlignment = Element.ALIGN_CENTER;
                                SPT2.BackgroundColor = new BaseColor(135, 0, 27);
                                SPT2.BorderWidth = 1F;
                                espTable.AddCell(SPT2);
                                PdfPCell U3 = new PdfPCell(new Phrase("U", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                                U3.HorizontalAlignment = Element.ALIGN_CENTER;
                                U3.BackgroundColor = new BaseColor(135, 0, 27);
                                U3.BorderWidth = 1F;
                                espTable.AddCell(U3);
                                PdfPCell SPT3 = new PdfPCell(new Phrase("T", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                                SPT3.HorizontalAlignment = Element.ALIGN_CENTER;
                                SPT3.BackgroundColor = new BaseColor(135, 0, 27);
                                SPT3.BorderWidth = 1F;
                                espTable.AddCell(SPT3);
                                var cour1 = "";
                                for (int a = 0; a < espT.Length - 1; a++)
                                {
                                    var esVal = espT[a].Split('|');
                                    var cour = "";
                                    if (esVal[0] == "" + grade + "TECH")
                                    {
                                        cour = "TECHNOLOGY";
                                    }
                                    else if (esVal[0] == "" + grade + "ART" || esVal[0] == "" + grade + "Art")
                                    {
                                        cour = "ART";
                                    }

                                    else if (esVal[0] == "" + grade + "MUS" || esVal[0] == "" + grade + "Mus")
                                    {
                                        cour = "MUSIC";
                                    }
                                    if (cour != cour1)
                                    {

                                        PdfPCell cou = new PdfPCell(new Phrase(cour, new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        cou.HorizontalAlignment = Element.ALIGN_LEFT;
                                        cou.BorderWidth = 1F;
                                        cou.Colspan = 6;
                                        cou.BorderColor = BaseColor.GRAY;
                                        espTable.AddCell(cou);
                                        PdfPCell tea = new PdfPCell(new Phrase(esVal[1], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.ITALIC, BaseColor.BLACK)));
                                        tea.HorizontalAlignment = Element.ALIGN_LEFT;
                                        tea.BorderWidth = 1F;
                                        tea.Colspan = 6;
                                        tea.BorderColor = BaseColor.GRAY;
                                        espTable.AddCell(tea);
                                        PdfPCell SPut1;
                                        if (esVal[2] != "")
                                        {
                                            SPut1 = new PdfPCell(new Phrase(esVal[2].Split('/')[1], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        }
                                        else
                                        {
                                            SPut1 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        }
                                        SPut1.HorizontalAlignment = Element.ALIGN_CENTER;
                                        SPut1.BorderWidth = 1F;
                                        SPut1.BorderColor = BaseColor.GRAY;
                                        espTable.AddCell(SPut1);
                                        PdfPCell SPut2;
                                        if (esVal[2] != "")
                                        {
                                            SPut2 = new PdfPCell(new Phrase(esVal[2].Split('/')[0], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        }
                                        else
                                        {
                                            SPut2 = new PdfPCell(new Phrase("", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        }
                                        SPut2.HorizontalAlignment = Element.ALIGN_CENTER;
                                        SPut2.BorderWidth = 1F;
                                        SPut2.BorderColor = BaseColor.GRAY;
                                        espTable.AddCell(SPut2);
                                        PdfPCell SPut3;
                                        if (esVal[3] != "")
                                        {
                                            SPut3 = new PdfPCell(new Phrase(esVal[3].Split('/')[1], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        }
                                        else
                                        {
                                            SPut3 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        }
                                        SPut3.HorizontalAlignment = Element.ALIGN_CENTER;
                                        SPut3.BorderWidth = 1F;
                                        SPut3.BorderColor = BaseColor.GRAY;
                                        espTable.AddCell(SPut3);
                                        PdfPCell SPut4;
                                        if (esVal[3] != "")
                                        {
                                            SPut4 = new PdfPCell(new Phrase(esVal[3].Split('/')[0], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        }
                                        else
                                        {
                                            SPut4 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        }

                                        SPut4.HorizontalAlignment = Element.ALIGN_CENTER;
                                        SPut4.BorderWidth = 1F;
                                        SPut4.BorderColor = BaseColor.GRAY;
                                        espTable.AddCell(SPut4);
                                        PdfPCell SPut5;
                                        if (esVal[4] != "")
                                        {
                                            SPut5 = new PdfPCell(new Phrase(esVal[4].Split('/')[1], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        }
                                        else
                                        {
                                            SPut5 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        }
                                        SPut5.HorizontalAlignment = Element.ALIGN_CENTER;
                                        SPut5.BorderWidth = 1F;
                                        SPut5.BorderColor = BaseColor.GRAY;
                                        espTable.AddCell(SPut5);
                                        PdfPCell SPut6;
                                        if (esVal[4] != "")
                                        {
                                            SPut6 = new PdfPCell(new Phrase(esVal[4].Split('/')[0], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        }
                                        else
                                        {
                                            SPut6 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        }
                                        SPut6.HorizontalAlignment = Element.ALIGN_CENTER;
                                        SPut6.BorderWidth = 1F;
                                        SPut6.BorderColor = BaseColor.GRAY;
                                        espTable.AddCell(SPut6);
                                        cour1 = cour;

                                    }
                                }
                                PdfPCell fo = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.WHITE)));
                                fo.HorizontalAlignment = Element.ALIGN_LEFT;
                                fo.Colspan = 18;
                                fo.Border = 0;
                                espTable.AddCell(fo);

                                PdfPCell ATTE = new PdfPCell(new Phrase("ATTENDANCE", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                                ATTE.HorizontalAlignment = Element.ALIGN_LEFT;
                                ATTE.Colspan = 6;
                                ATTE.Border = 1;
                                ATTE.BorderColor = BaseColor.GRAY;
                                ATTE.BackgroundColor = new BaseColor(135, 0, 27);
                                espTable.AddCell(ATTE);
                                PdfPCell A1 = new PdfPCell(new Phrase("T1", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                                A1.HorizontalAlignment = Element.ALIGN_CENTER;
                                A1.BackgroundColor = new BaseColor(135, 0, 27);
                                A1.BorderWidth = 1F;
                                A1.Border = 1;
                                A1.BorderColor = BaseColor.GRAY;
                                espTable.AddCell(A1);
                                PdfPCell A2 = new PdfPCell(new Phrase("T2", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                                A2.HorizontalAlignment = Element.ALIGN_CENTER;
                                A2.BackgroundColor = new BaseColor(135, 0, 27);
                                A2.BorderWidth = 1F;
                                A2.Border = 1;
                                A2.BorderColor = BaseColor.GRAY;
                                espTable.AddCell(A2);
                                PdfPCell A3 = new PdfPCell(new Phrase("T3", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                                A3.HorizontalAlignment = Element.ALIGN_CENTER;
                                A3.BackgroundColor = new BaseColor(135, 0, 27);
                                A3.BorderWidth = 1F;
                                A3.Border = 1;
                                A3.BorderColor = BaseColor.GRAY;
                                espTable.AddCell(A3);
                                PdfPCell spa5 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.WHITE)));
                                spa5.Colspan = 3;
                                spa5.Rowspan = 3;
                                spa5.Border = 0;
                                espTable.AddCell(spa5);
                                PdfPCell STK = new PdfPCell(new Phrase("STANDARD KEY", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                                STK.HorizontalAlignment = Element.ALIGN_LEFT;
                                STK.Colspan = 6;
                                STK.Border = 1;
                                STK.BorderWidth = 1F;
                                STK.BorderColor = BaseColor.GRAY;
                                STK.BackgroundColor = new BaseColor(135, 0, 27);
                                espTable.AddCell(STK);
                                PdfPCell ABSEN = new PdfPCell(new Phrase("Days Absent", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                ABSEN.HorizontalAlignment = Element.ALIGN_LEFT;
                                ABSEN.Colspan = 6;
                                ABSEN.BorderColor = BaseColor.GRAY;
                                espTable.AddCell(ABSEN);
                                if (ABTAR != "")
                                {
                                    if (ABTAR.Split('|')[0] != "")
                                    {
                                        PdfPCell abt = new PdfPCell(new Phrase(ABTAR.Split('|')[0].Split(',')[0], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        abt.HorizontalAlignment = Element.ALIGN_CENTER;
                                        abt.BorderColor = BaseColor.GRAY;
                                        espTable.AddCell(abt);
                                    }
                                    else
                                    {
                                        PdfPCell abt = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        abt.HorizontalAlignment = Element.ALIGN_CENTER;
                                        abt.BorderColor = BaseColor.GRAY;
                                        espTable.AddCell(abt);
                                    }

                                    if (ABTAR.Split('|')[1] != "")
                                    {

                                        PdfPCell abt1 = new PdfPCell(new Phrase(ABTAR.Split('|')[1].Split(',')[0], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        abt1.HorizontalAlignment = Element.ALIGN_CENTER;
                                        abt1.BorderColor = BaseColor.GRAY;
                                        espTable.AddCell(abt1);
                                    }
                                    else
                                    {
                                        PdfPCell abt1 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.WHITE)));
                                        abt1.HorizontalAlignment = Element.ALIGN_CENTER;
                                        abt1.BorderColor = BaseColor.GRAY;
                                        espTable.AddCell(abt1);
                                    }
                                    if (ABTAR.Split('|')[2] != "")
                                    {
                                        PdfPCell abt2 = new PdfPCell(new Phrase(ABTAR.Split('|')[2].Split(',')[0], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.WHITE)));
                                        abt2.HorizontalAlignment = Element.ALIGN_CENTER;
                                        abt2.BorderColor = BaseColor.GRAY;
                                        espTable.AddCell(abt2);
                                    }
                                    else
                                    {
                                        PdfPCell abt2 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.WHITE)));
                                        abt2.HorizontalAlignment = Element.ALIGN_CENTER;
                                        abt2.BorderColor = BaseColor.GRAY;
                                        espTable.AddCell(abt2);
                                    }
                                }
                                PdfPCell leg = new PdfPCell(new Phrase("U = Understands concepts and uses skills." + Environment.NewLine + "T = Follows Tribes® Agreements.", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                leg.HorizontalAlignment = Element.ALIGN_LEFT;
                                leg.Colspan = 6;
                                leg.Rowspan = 2;
                                leg.Border = 1;
                                leg.BorderColor = BaseColor.GRAY;
                                espTable.AddCell(leg);

                                PdfPCell TARD = new PdfPCell(new Phrase("Days Tardy", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                TARD.HorizontalAlignment = Element.ALIGN_LEFT;
                                TARD.BorderColor = BaseColor.GRAY;
                                TARD.Colspan = 6;
                                espTable.AddCell(TARD);
                                if (ABTAR != "")
                                {
                                    if (ABTAR.Split('|')[0] != "")
                                    {
                                        PdfPCell tard1 = new PdfPCell(new Phrase(ABTAR.Split('|')[0].Split(',')[1], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        tard1.HorizontalAlignment = Element.ALIGN_CENTER;
                                        tard1.BorderColor = BaseColor.GRAY;
                                        espTable.AddCell(tard1);
                                    }
                                    else
                                    {
                                        PdfPCell tard1 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        tard1.HorizontalAlignment = Element.ALIGN_CENTER;
                                        tard1.BorderColor = BaseColor.GRAY;
                                        espTable.AddCell(tard1);
                                    }
                                    if (ABTAR.Split('|')[1] != "")
                                    {
                                        PdfPCell tard2 = new PdfPCell(new Phrase(ABTAR.Split('|')[1].Split(',')[1], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        tard2.HorizontalAlignment = Element.ALIGN_CENTER;
                                        tard2.BorderColor = BaseColor.GRAY;
                                        espTable.AddCell(tard2);
                                    }
                                    else
                                    {
                                        PdfPCell tard2 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.WHITE)));
                                        tard2.HorizontalAlignment = Element.ALIGN_CENTER;
                                        tard2.BorderColor = BaseColor.GRAY;
                                        espTable.AddCell(tard2);
                                    }
                                    if (ABTAR.Split('|')[2] != "")
                                    {
                                        PdfPCell tard4 = new PdfPCell(new Phrase(ABTAR.Split('|')[2].Split(',')[1], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        tard4.HorizontalAlignment = Element.ALIGN_CENTER;
                                        tard4.BorderColor = BaseColor.GRAY;
                                        espTable.AddCell(tard4);
                                    }
                                    else
                                    {
                                        PdfPCell tard4 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.WHITE)));
                                        tard4.HorizontalAlignment = Element.ALIGN_CENTER;
                                        tard4.BorderColor = BaseColor.GRAY;
                                        espTable.AddCell(tard4);
                                    }

                                }
                            }
                            else
                            {
                                PdfPCell ep21 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                                ep21.HorizontalAlignment = Element.ALIGN_LEFT;
                                ep21.Colspan = 18;
                                ep21.Border = 0;
                                espTable.AddCell(ep21);
                                PdfPCell ATTE = new PdfPCell(new Phrase("ATTENDANCE", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                                ATTE.HorizontalAlignment = Element.ALIGN_LEFT;
                                ATTE.Colspan = 15;
                                ATTE.Border = 1;
                                ATTE.BorderColor = BaseColor.GRAY;
                                ATTE.BackgroundColor = new BaseColor(135, 0, 27);
                                espTable.AddCell(ATTE);
                                PdfPCell A1 = new PdfPCell(new Phrase("T1", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                                A1.HorizontalAlignment = Element.ALIGN_CENTER;
                                A1.BackgroundColor = new BaseColor(135, 0, 27);
                                A1.BorderWidth = 1F;
                                A1.Border = 1;
                                A1.BorderColor = BaseColor.GRAY;
                                espTable.AddCell(A1);
                                PdfPCell A2 = new PdfPCell(new Phrase("T2", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                                A2.HorizontalAlignment = Element.ALIGN_CENTER;
                                A2.BackgroundColor = new BaseColor(135, 0, 27);
                                A2.BorderWidth = 1F;
                                A2.Border = 1;
                                A2.BorderColor = BaseColor.GRAY;
                                espTable.AddCell(A2);
                                PdfPCell A3 = new PdfPCell(new Phrase("T3", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                                A3.HorizontalAlignment = Element.ALIGN_CENTER;
                                A3.BackgroundColor = new BaseColor(135, 0, 27);
                                A3.BorderWidth = 1F;
                                A3.Border = 1;
                                A3.BorderColor = BaseColor.GRAY;
                                espTable.AddCell(A3);
                                PdfPCell ABSEN = new PdfPCell(new Phrase("Days Absent", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                ABSEN.HorizontalAlignment = Element.ALIGN_LEFT;
                                ABSEN.Colspan = 15;
                                ABSEN.BorderColor = BaseColor.GRAY;
                                espTable.AddCell(ABSEN);
                                if (ABTAR != "")

                                {
                                    if (ABTAR.Split('|')[0] != "")
                                    {

                                        PdfPCell abt = new PdfPCell(new Phrase(ABTAR.Split('|')[0].Split(',')[0], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        abt.HorizontalAlignment = Element.ALIGN_CENTER;
                                        abt.BorderColor = BaseColor.GRAY;
                                        espTable.AddCell(abt);
                                    }
                                    else
                                    {
                                        PdfPCell abt = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        abt.HorizontalAlignment = Element.ALIGN_CENTER;
                                        abt.BorderColor = BaseColor.GRAY;
                                        espTable.AddCell(abt);
                                    }
                                    if (ABTAR.Split('|')[1] != "")
                                    {
                                        PdfPCell abt1 = new PdfPCell(new Phrase(ABTAR.Split('|')[1].Split(',')[0], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        abt1.HorizontalAlignment = Element.ALIGN_CENTER;
                                        abt1.BorderColor = BaseColor.GRAY;
                                        espTable.AddCell(abt1);
                                    }else
                                    {
                                        PdfPCell abt1 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.WHITE)));
                                        abt1.HorizontalAlignment = Element.ALIGN_CENTER;
                                        abt1.BorderColor = BaseColor.GRAY;
                                        espTable.AddCell(abt1);
                                    }
                                    if (ABTAR.Split('|')[2] != "")
                                    {
                                        PdfPCell abt2 = new PdfPCell(new Phrase(ABTAR.Split('|')[2].Split(',')[0], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        abt2.HorizontalAlignment = Element.ALIGN_CENTER;
                                        abt2.BorderColor = BaseColor.GRAY;
                                        espTable.AddCell(abt2);
                                    }else
                                    {
                                        PdfPCell abt2 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.WHITE)));
                                        abt2.HorizontalAlignment = Element.ALIGN_CENTER;
                                        abt2.BorderColor = BaseColor.GRAY;
                                        espTable.AddCell(abt2);
                                    }

                                }
                                else
                                {
                                    PdfPCell tard1 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.WHITE)));
                                    tard1.HorizontalAlignment = Element.ALIGN_CENTER;
                                    tard1.BorderColor = BaseColor.GRAY;
                                    espTable.AddCell(tard1);
                                    PdfPCell tard2 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.WHITE)));
                                    tard2.HorizontalAlignment = Element.ALIGN_CENTER;
                                    tard2.BorderColor = BaseColor.GRAY;
                                    espTable.AddCell(tard2);
                                    PdfPCell tard3 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.WHITE)));
                                    tard3.HorizontalAlignment = Element.ALIGN_CENTER;
                                    tard3.BorderColor = BaseColor.GRAY;
                                    espTable.AddCell(tard3);
                                }
                                PdfPCell TARD = new PdfPCell(new Phrase("Days Tardy", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                TARD.HorizontalAlignment = Element.ALIGN_LEFT;
                                TARD.BorderColor = BaseColor.GRAY;
                                TARD.Colspan = 15;
                                espTable.AddCell(TARD);
                                if (ABTAR != "")
                                {
                                    if (ABTAR.Split('|')[0] != "")
                                    {
                                        PdfPCell tard1 = new PdfPCell(new Phrase(ABTAR.Split('|')[0].Split(',')[1], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        tard1.HorizontalAlignment = Element.ALIGN_CENTER;
                                        tard1.BorderColor = BaseColor.GRAY;
                                        espTable.AddCell(tard1);
                                    }
                                    else
                                    {
                                        PdfPCell tard1 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        tard1.HorizontalAlignment = Element.ALIGN_CENTER;
                                        tard1.BorderColor = BaseColor.GRAY;
                                        espTable.AddCell(tard1);
                                    }
                                    if (ABTAR.Split('|')[1] != "")
                                    {
                                        PdfPCell tard2 = new PdfPCell(new Phrase(ABTAR.Split('|')[1].Split(',')[1], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        tard2.HorizontalAlignment = Element.ALIGN_CENTER;
                                        tard2.BorderColor = BaseColor.GRAY;
                                        espTable.AddCell(tard2);
                                    }else
                                    {
                                        PdfPCell tard2 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.WHITE)));
                                        tard2.HorizontalAlignment = Element.ALIGN_CENTER;
                                        tard2.BorderColor = BaseColor.GRAY;
                                        espTable.AddCell(tard2);
                                    }
                                    if (ABTAR.Split('|')[2] != "")
                                    {
                                        PdfPCell tard4 = new PdfPCell(new Phrase(ABTAR.Split('|')[2].Split(',')[1], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        tard4.HorizontalAlignment = Element.ALIGN_CENTER;
                                        tard4.BorderColor = BaseColor.GRAY;
                                        espTable.AddCell(tard4);
                                    }else
                                    {
                                        PdfPCell tard4 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.WHITE)));
                                        tard4.HorizontalAlignment = Element.ALIGN_CENTER;
                                        tard4.BorderColor = BaseColor.GRAY;
                                        espTable.AddCell(tard4);
                                    }
                                }else
                                {
                                    PdfPCell tard1 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.WHITE)));
                                    tard1.HorizontalAlignment = Element.ALIGN_CENTER;
                                    tard1.BorderColor = BaseColor.GRAY;
                                    espTable.AddCell(tard1);
                                    PdfPCell tard2 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.WHITE)));
                                    tard2.HorizontalAlignment = Element.ALIGN_CENTER;
                                    tard2.BorderColor = BaseColor.GRAY;
                                    espTable.AddCell(tard2);
                                    PdfPCell tard3 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.WHITE)));
                                    tard3.HorizontalAlignment = Element.ALIGN_CENTER;
                                    tard3.BorderColor = BaseColor.GRAY;
                                    espTable.AddCell(tard3);
                                }
                            }
                            PdfPCell ap6 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.WHITE)));
                            ap6.HorizontalAlignment = Element.ALIGN_CENTER;
                            ap6.Colspan = 18;
                            ap6.Border = 0;
                            espTable.AddCell(ap6);
                            PdfPCell comm1 = new PdfPCell(new Phrase("T1 Comment ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.BLACK)));
                            comm1.HorizontalAlignment = Element.ALIGN_LEFT;
                            comm1.Colspan = 18;
                            comm1.Border = 0;
                            comm1.BorderWidthBottom = 1;
                            comm1.BorderColorBottom = BaseColor.LIGHT_GRAY;
                            espTable.AddCell(comm1);

                            PdfPCell comms = new PdfPCell(new Phrase(comm, new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                            comms.HorizontalAlignment = Element.ALIGN_LEFT;
                            comms.Colspan = 18;
                            comms.Border = 0;
                            espTable.AddCell(comms);



                            documento.Add(HeadT);
                            documento.Add(legendTable);
                            documento.Add(GradeTable);
                            documento.Add(espTable);

                            //Process prc = new System.Diagnostics.Process();
                            //prc.StartInfo.FileName = fileName;
                            //prc.Start();
                        }
                        else
                        {
                            con.Close();

                        }
                        documento.NewPage();
                    }

                    con.Close();

                }
                else
                {
                    if (con.State == ConnectionState.Closed)
                    {
                        con.Open();
                    }
                    fname = "ES_" + grade + "Gr_ReportCard_" + DateTime.Now.DayOfYear + DateTime.Now.Month + DateTime.Now.Day + DateTime.Now.Millisecond + ".pdf";
                    fileName = HttpContext.Current.Server.MapPath("~/RepoFiles/" + fname);
                    PdfWriter.GetInstance(documento, new FileStream(fileName, FileMode.Create));
                    documento.Open();

                    string T1DATA = string.Empty;
                    string T1AD = string.Empty;
                    string T1ESP = string.Empty;

                    if (con.State == ConnectionState.Closed)
                    {
                        con.Open();
                    }

                    sql = " CREATE OR REPLACE VIEW ES_VISTA";
                    sql += " AS WITH X AS(";
                    sql += " SELECT IDENTIFIER, STDID, STDDESC, stcourse, STUDENT, STUDENT_NUMBER, STDCID, STID, GRADE_LEVEL, SUBJECTAREA FROM(";
                    sql += " SELECT ST.STANDARDID STDID, ST.IDENTIFIER, TO_CHAR(TRANSIENTCOURSELIST) stcourse, ST.SUBJECTAREA, ST.NAME STDDESC FROM STANDARD ST";
                    if (grade == "2")
                    {
                        sql += " WHERE TO_CHAR(TRANSIENTCOURSELIST)  IN('" + grade + "Ma', '" + grade + "SS', '" + grade + "LA', '" + grade + "SLA', '" + grade + "HR', '" + grade + "SC', '" + grade + "PE')";
                    }
                    else if (grade == "PK")
                    {
                        sql += " WHERE TO_CHAR(TRANSIENTCOURSELIST)  IN('" + grade + "MA','" + grade + "LA', '" + grade + "SLA', '" + grade + "HR', '" + grade + "SC," + grade + "SS', '" + grade + "PE')";
                    }
                    else
                    {
                        sql += " WHERE TO_CHAR(TRANSIENTCOURSELIST)  IN('" + grade + "MA', '" + grade + "SS', '" + grade + "LA', '" + grade + "SLA', '" + grade + "HR', '" + grade + "SC', '" + grade + "PE')";
                    }
                    sql += " AND ST.YEARID = 28 AND ST.STANDARDID NOT IN(17299, 17293, 17923)AND isassignmentallowed = 1 AND ISACTIVE = 1)";
                    sql += " CROSS JOIN";
                    sql += " (";
                    sql += " SELECT FIRST_NAME || ' ' || LAST_NAME STUDENT, STUDENT_NUMBER, STUDENTS.DCID AS STDCID, ID AS STID, GRADE_LEVEL FROM STUDENTS";
                    if (grade == "K")
                    {
                        sql += " WHERE GRADE_LEVEL=0 AND ENROLL_STATUS = 0 AND STUDENT_NUMBER =" + stnum + ")";
                    }
                    else if (grade == "PK")
                    {
                        sql += " WHERE GRADE_LEVEL=-1 AND ENROLL_STATUS = 0 AND STUDENT_NUMBER =" + stnum + ")";
                    }
                    else
                    {
                        sql += " WHERE GRADE_LEVEL='" + grade + "' AND ENROLL_STATUS = 0 AND STUDENT_NUMBER =" + stnum + ")";
                    }
                    sql += " )";
                    sql += " SELECT IDENTIFIER,STDID,STDDESC,STCOURSE,STUDENT,STUDENT_NUMBER,STDCID,STID,GRADE_LEVEL,SUBJECTAREA FROM X";

                    OracleCommand cmdV1 = new OracleCommand(sql, con);
                    cmdV1.ExecuteNonQuery();


                    sql = " WITH MQUERY AS(SELECT IDENTIFIER, STDID, STDDESC, STCOURSE, STUDENT, STUDENT_NUMBER, STDCID, STID, GRADE_LEVEL,";
                    sql += " SUBJECTAREA, SG.STORECODE, SG.STANDARDGRADE, T.LASTFIRST TEACHER";
                    sql += " FROM ES_VISTA";
                    sql += " LEFT JOIN STANDARDGRADESECTION SG ON STDCID = SG.STUDENTSDCID AND STDID = SG.STANDARDID AND SG.STANDARDID IS NOT NULL AND SG.STORECODE IN('T1', 'T2', 'T3') AND SG.STANDARDGRADE<>'--'";
                    sql += " LEFT JOIN CC CO ON STID = CO.STUDENTID AND STCOURSE = CO.COURSE_NUMBER  AND CO.ORIGSECTIONID = 0";
                    sql += " LEFT JOIN TEACHERS T ON CO.TEACHERID = T.ID)";
                    sql += " SELECT DISTINCT IDENTIFIER,STID,STUDENT_NUMBER,STUDENT,GRADE_LEVEL,GRADE_LEVEL,TEACHER,STCOURSE,SUBJECTAREA,STDDESC";
                    sql += " ,(SELECT DISTINCT Y.STANDARDGRADE FROM MQUERY y WHERE y.IDENTIFIER = M.IDENTIFIER AND Y.STORECODE = 'T1') T1";
                    sql += " ,(SELECT DISTINCT Y.STANDARDGRADE FROM MQUERY y WHERE y.IDENTIFIER = M.IDENTIFIER AND Y.STORECODE = 'T2') T2";
                    sql += " ,(SELECT DISTINCT Y.STANDARDGRADE FROM MQUERY y WHERE y.IDENTIFIER = M.IDENTIFIER AND Y.STORECODE = 'T3') T3";
                    sql += " FROM MQUERY M";

                    sql += " ORDER BY CASE";
                    sql += " WHEN STCOURSE LIKE '%" + grade + "HR%' THEN 1";
                    sql += " WHEN STCOURSE LIKE '%" + grade + "LA%' THEN 2";
                    if (grade == "2")
                    {
                        sql += " WHEN STCOURSE LIKE '%" + grade + "Ma%' THEN 3";
                    }
                    else
                    {
                        sql += " WHEN STCOURSE LIKE '%" + grade + "MA%' THEN 3";
                    }
                    sql += " WHEN STCOURSE LIKE '%" + grade + "SLA%' THEN 4";
                    sql += " WHEN STCOURSE LIKE '%" + grade + "SS%' THEN 5";
                    sql += " WHEN STCOURSE LIKE '%" + grade + "SC%' THEN 6";
                    sql += " WHEN STCOURSE LIKE '%" + grade + "PE%' THEN 7";
                    sql += " END,IDENTIFIER ASC";


                   

                    OracleCommand cmd1 = new OracleCommand(sql, con);
                    OracleDataReader odr1 = cmd1.ExecuteReader();
                    while (odr1.Read())
                    {
                        T1DATA += odr1["STUDENT_NUMBER"].ToString() + '|';
                        T1DATA += odr1["STUDENT"].ToString() + '|';
                        T1DATA += odr1["GRADE_LEVEL"].ToString() + '|';
                        T1DATA += odr1["TEACHER"].ToString() + '|';
                        T1DATA += odr1["STCOURSE"].ToString() + '|';
                        T1DATA += odr1["SUBJECTAREA"].ToString() + '|';
                        T1DATA += odr1["STDDESC"].ToString() + '|';
                        T1DATA += odr1["T1"].ToString() + '|';
                        T1DATA += odr1["T2"].ToString() + '|';
                        T1DATA += odr1["T3"].ToString() + '|';
                        T1DATA += odr1["STID"].ToString() + '|';
                        T1DATA += odr1["IDENTIFIER"].ToString() + '^';

                    }

                    //sql += " CREATE OR REPLACE VIEW ES_VISTA_ESP";
                    //sql += " AS WITH X AS(";
                    //sql += " SELECT IDENTIFIER, STDID, STDDESC, stcourse, STUDENT, STUDENT_NUMBER, STDCID, STID, GRADE_LEVEL, SUBJECTAREA FROM(";
                    //sql += " SELECT ST.STANDARDID STDID, ST.IDENTIFIER, TO_CHAR(TRANSIENTCOURSELIST) stcourse, ST.SUBJECTAREA, ST.NAME STDDESC FROM STANDARD ST";
                    //if (grade == "2")
                    //{
                    //    sql += " WHERE TO_CHAR(TRANSIENTCOURSELIST)  IN('" + grade + "TECH', '" + grade + "Art', '" + grade + "Mus')";
                    //}
                    //else { 
                    //    sql += " WHERE TO_CHAR(TRANSIENTCOURSELIST)  IN('" + grade + "TECH', '" + grade + "ART', '" + grade + "MUS')";
                    //{
                    //sql += " AND ST.YEARID = 28 AND ST.STANDARDID NOT IN(17299, 17293, 17923)AND isassignmentallowed = 1 AND ISACTIVE = 1)";
                    //sql += " CROSS JOIN";
                    //sql += " (";
                    //sql += " SELECT FIRST_NAME || ' ' || LAST_NAME STUDENT, STUDENT_NUMBER, STUDENTS.DCID AS STDCID, ID AS STID, GRADE_LEVEL FROM STUDENTS";
                    //if (grade == "K")
                    //{
                    //    sql += " WHERE GRADE_LEVEL=0 AND ENROLL_STATUS = 0 AND STUDENT_NUMBER =" + stnum + ")";
                    //}
                    //else if (grade == "PK")
                    //{
                    //    sql += " WHERE GRADE_LEVEL=-1 AND ENROLL_STATUS = 0 AND STUDENT_NUMBER =" + stnum + ")";
                    //}
                    //else
                    //{
                    //    sql += " WHERE GRADE_LEVEL='" + grade + "' AND ENROLL_STATUS = 0 AND STUDENT_NUMBER =" + stnum + ")";
                    //}
                    //sql += " )";
                    //sql += " SELECT IDENTIFIER, STDID, STDDESC, STCOURSE, STUDENT, STUDENT_NUMBER, STDCID, STID, GRADE_LEVEL, SUBJECTAREA FROM X;";

                    //OracleCommand cmdV2 = new OracleCommand(sql, con);
                    //cmdV2.ExecuteNonQuery();

                    //sql += " SELECT TEACHER, STCOURSE, LISTAGG(T1,',') WITHIN GROUP (ORDER BY STDDESC) T1,LISTAGG(T2, ',') WITHIN GROUP (ORDER BY STDDESC)T2,LISTAGG(T3, ',') WITHIN GROUP (ORDER BY STDDESC)T3 FROM (";
                    //sql += " WITH MQUERY AS(SELECT IDENTIFIER, STDID, STDDESC, STCOURSE, STUDENT, STUDENT_NUMBER, STDCID, STID, GRADE_LEVEL,";
                    //sql += " SUBJECTAREA, SG.STORECODE, SG.STANDARDGRADE, T.LASTFIRST TEACHER";
                    //sql += " FROM ES_VISTA_ESP";
                    //sql += " LEFT JOIN STANDARDGRADESECTION SG ON STDCID = SG.STUDENTSDCID AND sg.yearid = 28 and STDID = SG.STANDARDID AND STANDARDGRADE <> '--'";
                    //sql += "  AND SG.STANDARDID IS NOT NULL AND SG.STORECODE IN('T1', 'T2', 'T3')";
                    //sql += " LEFT JOIN CC CO ON STID = CO.STUDENTID AND STCOURSE = CO.COURSE_NUMBER  AND CO.ORIGSECTIONID = 0";
                    //sql += " LEFT JOIN TEACHERS T ON CO.TEACHERID = T.ID)";
                    //sql += " SELECT DISTINCT TEACHER, STCOURSE, STDDESC";
                    //sql += " , (SELECT distinct Y.STANDARDGRADE FROM MQUERY y WHERE y.IDENTIFIER = M.IDENTIFIER AND Y.STORECODE = 'T1') T1";
                    //sql += " ,(SELECT distinct Y.STANDARDGRADE FROM MQUERY y WHERE y.IDENTIFIER = M.IDENTIFIER AND Y.STORECODE = 'T2') T2";
                    //sql += " ,(SELECT distinct Y.STANDARDGRADE FROM MQUERY y WHERE y.IDENTIFIER = M.IDENTIFIER AND Y.STORECODE = 'T3') T3";
                    //sql += " FROM MQUERY M";
                    //sql += " )";
                    //sql += " GROUP BY TEACHER,STCOURSE";
                    //sql += " ORDER BY CASE";
                    //sql += " WHEN STCOURSE LIKE '%" + grade + "TECH%' THEN 1";
                    //if (grade == "2")
                    //{
                    //    sql += " WHEN STCOURSE LIKE '%" + grade + "Art%' THEN 2";
                    //    sql += " WHEN STCOURSE LIKE '%" + grade + "Mus%' THEN 3";
                    //}
                    //else
                    //{
                    //    sql += " WHEN STCOURSE LIKE '%" + grade + "ART%' THEN 2";
                    //    sql += " WHEN STCOURSE LIKE '%" + grade + "MUS%' THEN 3";
                    //}
                    //    sql += " END";


                    sql = "WITH X AS(";
                    sql += " SELECT IDENTIFIER, STANDARDID, STDDESC, stcourse, LASTFIRST, STUDENT_NUMBER, STDCID, STID, GRADE_LEVEL, SUBJECTAREA FROM(";
                    sql += " SELECT ST.STANDARDID, ST.IDENTIFIER, TO_CHAR(TRANSIENTCOURSELIST) stcourse, ST.SUBJECTAREA, ST.NAME STDDESC FROM STANDARD ST";
                    if (grade == "2")
                    {
                        sql += " WHERE TO_CHAR(TRANSIENTCOURSELIST)  IN('" + grade + "TECH', '" + grade + "Art', '" + grade + "Mus')";
                    }
                    else
                    {
                        sql += " WHERE TO_CHAR(TRANSIENTCOURSELIST)  IN('" + grade + "TECH', '" + grade + "ART', '" + grade + "MUS')";
                    }
                    sql += " AND ST.YEARID = 28 AND STANDARDID NOT IN(17299, 17293)AND isassignmentallowed = 1)";
                    sql += " CROSS JOIN";
                    sql += " (";
                    sql += " SELECT LASTFIRST, STUDENT_NUMBER, STUDENTS.DCID AS STDCID, ID AS STID, GRADE_LEVEL FROM STUDENTS";
                    if (grade == "K")
                    {
                        sql += " WHERE GRADE_LEVEL=0 AND ENROLL_STATUS = 0 AND STUDENT_NUMBER =" + stnum + ")";
                    }
                    else if (grade == "PK")
                    {
                        sql += " WHERE GRADE_LEVEL=-1 AND ENROLL_STATUS = 0 AND STUDENT_NUMBER =" + stnum + ")";
                    }
                    else
                    {
                        sql += " WHERE GRADE_LEVEL='" + grade + "' AND ENROLL_STATUS = 0 AND STUDENT_NUMBER =" + stnum + ")";
                    }
                    sql += " )";
                    sql += " SELECT STUDENT, TEACHER, STCOURSE,(CASE WHEN STORECODE = 'T1' THEN T_GRADE || '/' || U_GRADE ELSE NULL END)T1";
                    sql += " ,(CASE WHEN STORECODE = 'T2' THEN T_GRADE|| '/' || U_GRADE ELSE NULL END)T2";
                    sql += " ,(CASE WHEN STORECODE = 'T3' THEN T_GRADE|| '/' || U_GRADE ELSE NULL END)T3";
                    sql += " FROM(SELECT X.STDDESC, X.STID, X.STUDENT_NUMBER, X.LASTFIRST STUDENT, T.LASTFIRST TEACHER, X.STCOURSE, SG.STANDARDGRADE, SG.STORECODE";
                    sql += " FROM X";
                    sql += " LEFT JOIN STANDARDGRADESECTION SG ON X.STDCID = SG.STUDENTSDCID AND X.STANDARDID = SG.STANDARDID AND SG.STANDARDID IS NOT NULL AND SG.STORECODE IN('T1') AND SG.STANDARDGRADE<>'--'";
                    sql += " LEFT JOIN CC CO ON X.STID = CO.STUDENTID AND X.STCOURSE = CO.COURSE_NUMBER  AND CO.ORIGSECTIONID = 0";
                    sql += " LEFT JOIN TEACHERS T ON CO.TEACHERID = T.ID";
                    sql += " )";
                    sql += " PIVOT(MAX(standardgrade) AS grade FOR(STDDESC) IN('Understands concepts and uses skills.' AS U,'Follows Tribes® Agreements.' AS T))";
                    sql += " ORDER BY STUDENT,STCOURSE DESC";

                    OracleCommand cmd2 = new OracleCommand(sql, con);
                    OracleDataReader odr2 = cmd2.ExecuteReader();
                    while (odr2.Read())
                    {
                        T1ESP += odr2["STCOURSE"].ToString() + '|';
                        T1ESP += odr2["TEACHER"].ToString() + '|';
                        T1ESP += odr2["T1"].ToString() + '|';
                        T1ESP += odr2["T2"].ToString() + '|';
                        T1ESP += odr2["T3"].ToString() + '^';
                    }

                    sql = " SELECT LISTAGG(T1,',') WITHIN GROUP (ORDER BY STUDENT) T1,LISTAGG(T2, ',') WITHIN GROUP (ORDER BY STUDENT) T2,LISTAGG(T3, ',') WITHIN GROUP (ORDER BY STUDENT) T3 FROM (SELECT STUDENT,";
                    sql += " (CASE WHEN ABBRE = 'T1' THEN COUNT(ABSENCE) || ',' || COUNT(TARDI) END)T1,(CASE WHEN ABBRE = 'T2' THEN COUNT(ABSENCE) || ',' || COUNT(TARDI) END)T2";
                    sql += " ,(CASE WHEN ABBRE = 'T3' THEN COUNT(ABSENCE) || ',' || COUNT(TARDI) END)T3 FROM (SELECT DISTINCT S.LASTFIRST STUDENT,";
                    sql += " AC.ATT_CODE, (CASE WHEN(AC.ATT_CODE = 'EA' OR AC.ATT_CODE = 'UA') THEN AC.PRESENCE_STATUS_CD END) ABSENCE,";
                    sql += " (CASE  WHEN(AC.ATT_CODE = 'ET' OR AC.ATT_CODE = 'UT') THEN AC.PRESENCE_STATUS_CD END) TARDI, AT.ATT_DATE,T.ABBREVIATION ABBRE FROM ATTENDANCE AT";
                    sql += " LEFT JOIN STUDENTS S ON AT.STUDENTID = S.ID";
                    sql += " LEFT JOIN ATTENDANCE_CODE AC ON AT.ATTENDANCE_CODEID = AC.ID";
                    sql += " LEFT JOIN TERMS T ON AT.YEARID = T.YEARID";
                    sql += " WHERE AT.YEARID = 28 AND AT.ATT_DATE BETWEEN (T.FIRSTDAY)AND(T.LASTDAY) AND T.ABBREVIATION IN('T1') AND S.STUDENT_NUMBER = " + stnum + "";
                    sql += " )";
                    sql += " GROUP BY STUDENT,ABBRE)";
                    sql += " GROUP BY STUDENT";

                    //sql = " SELECT STUDENT, COUNT(ABSENCE) ABSE, COUNT(TARDI) TARD FROM(SELECT DISTINCT S.LASTFIRST STUDENT,";
                    //sql += " AC.ATT_CODE, (CASE WHEN(AC.ATT_CODE = 'EA' OR AC.ATT_CODE = 'UA') THEN AC.PRESENCE_STATUS_CD END) ABSENCE,";
                    //sql += " (CASE  WHEN(AC.ATT_CODE = 'ET' OR AC.ATT_CODE = 'UT' ) THEN AC.PRESENCE_STATUS_CD END) TARDI, AT.ATT_DATE FROM ATTENDANCE AT";
                    //sql += " LEFT JOIN STUDENTS S ON AT.STUDENTID = S.ID";
                    //sql += " LEFT JOIN ATTENDANCE_CODE AC ON AT.ATTENDANCE_CODEID = AC.ID";
                    //sql += " WHERE AT.YEARID = 28 AND AT.ATT_DATE <= CURRENT_DATE AND S.STUDENT_NUMBER = " + stnum + "";
                    //sql += " )";
                    //sql += " GROUP BY STUDENT";

                    


                    string ABTAR = string.Empty;
                    OracleCommand cmd3 = new OracleCommand(sql, con);
                    OracleDataReader odr3 = cmd3.ExecuteReader();
                    while (odr3.Read())
                    {
                        ABTAR += odr3["T1"].ToString() + '|';
                        ABTAR += odr3["T2"].ToString() + '|';
                        ABTAR += odr3["T3"].ToString() + '|';
                    }

                    sql = " SELECT S.LASTFIRST STUDENT, ST.COMMENTVALUE FROM STANDARDGRADESECTIONCOMMENT ST";
                    sql += " LEFT JOIN STUDENTS S ON ST.STUDENTSDCID = S.DCID";
                    sql += " LEFT JOIN STANDARDGRADESECTION SG ON ST.STANDARDGRADESECTIONID = SG.STANDARDGRADESECTIONID";
                    if (grade == "K")
                    {
                        sql += " WHERE ST.YEARID = 28 AND S.GRADE_LEVEL='0' AND SG.STORECODE='T1' AND S.STUDENT_NUMBER = " + stnum + "";
                    }
                    else if (grade == "PK")
                    {
                        sql += " WHERE ST.YEARID = 28 AND S.GRADE_LEVEL='-1' AND SG.STORECODE='T1' AND S.STUDENT_NUMBER = " + stnum + "";
                    }
                    else
                    {
                        sql += " WHERE ST.YEARID = 28 AND S.GRADE_LEVEL='" + grade + "' AND SG.STORECODE='T1' AND S.STUDENT_NUMBER = " + stnum + "";
                    }

                    string comm = string.Empty;
                    OracleCommand cmd4 = new OracleCommand(sql, con);
                    OracleDataReader odr4 = cmd4.ExecuteReader();
                    while (odr4.Read())
                    {
                        comm = odr4["COMMENTVALUE"].ToString();
                    }

                    if (T1DATA != "")
                    {
                        

                        var stTable = T1DATA.Split('^');
                        var HT = "";
                        var std = "";
                        var stn = "";
                        var stgd = "";
                        var stid = "";
                        for (int i = 0; i < stTable.Length; i++)
                        {
                            var hr = stTable[i].Split('|');
                            if (hr[11].Split('.')[2] == "HWr" )
                            {
                                HT = hr[3];
                                std = hr[0];
                                stn = hr[1];
                                stgd = hr[2];
                                stid = hr[10];
                                break;
                            }else if (hr[4] == "PKHR")
                            {
                                HT = hr[3];
                                std = hr[0];
                                stn = hr[1];
                                stgd = hr[2];
                                stid = hr[10];
                                break;
                            }
                            
                        }



                        iTextSharp.text.Image Imagen = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("~/img/WLOGO.jpg"));
                        Imagen.ScalePercent(2.5f);
                       
                        iTextSharp.text.Image foto = iTextSharp.text.Image.GetInstance("file://cms03pws/e$/program%20files/powerschool/data/picture/student/"+ stid.Substring(stid.Length - Math.Min(2, stid.Length)) +"/"+stid+ "/ph.jpeg");
                       foto.ScalePercent(27f);
                    

                        PdfPTable HeadT = new PdfPTable(16);
                        HeadT.HorizontalAlignment = Element.ALIGN_CENTER;
                        HeadT.WidthPercentage = 100;

                        PdfPCell logo = new PdfPCell(Imagen);
                        logo.Colspan = 9;
                        logo.Border = 0;
                        logo.HorizontalAlignment = Element.ALIGN_LEFT;
                        logo.Rowspan = 3;
                        logo.Padding = 3;
                        HeadT.AddCell(logo);


                        PdfPCell HS = new PdfPCell(new Phrase("ELEMENTARY SCHOOL Report Card", new Font(Font.FontFamily.HELVETICA, 14, Font.BOLD, new BaseColor(135, 0, 27))));
                        HS.HorizontalAlignment = Element.ALIGN_BOTTOM;
                        HS.Colspan = 7;
                        HS.Border = 0;
                        HeadT.AddCell(HS);

                        PdfPCell SQ1 = new PdfPCell(new Phrase("School Year 2018-19 Trimester 1", new Font(Font.FontFamily.HELVETICA, 14, Font.BOLD, BaseColor.BLACK)));
                        SQ1.HorizontalAlignment = Element.ALIGN_BOTTOM;
                        SQ1.Colspan = 7;
                        SQ1.Border = 0;
                        HeadT.AddCell(SQ1);

                        PdfPCell Pub = new PdfPCell(new Phrase("Published " + DateTime.Now.ToString("MMMM dd, yyyy"), new Font(Font.FontFamily.HELVETICA, 12, Font.ITALIC, BaseColor.BLACK)));
                        Pub.Colspan = 7;
                        Pub.HorizontalAlignment = Element.ALIGN_BOTTOM;
                        Pub.Rowspan = 2;
                        Pub.Border = 0;
                        HeadT.AddCell(Pub);

                        PdfPCell bar1 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                        bar1.HorizontalAlignment = Element.ALIGN_LEFT;
                        bar1.Border = 0;
                        bar1.Colspan = 16;
                        bar1.BackgroundColor = new BaseColor(135, 0, 27);
                        HeadT.AddCell(bar1);

                        PdfPCell stinfo = new PdfPCell(new Phrase("Student Name: " + stn, new Font(Font.FontFamily.HELVETICA, 11, Font.BOLD, BaseColor.BLACK)));
                        stinfo.HorizontalAlignment = Element.ALIGN_LEFT;
                        stinfo.Border = 0;
                        stinfo.Colspan = 7;
                        stinfo.PaddingTop = 5;
                        HeadT.AddCell(stinfo);

                        PdfPCell stfoto = new PdfPCell(foto);
                        stfoto.HorizontalAlignment = Element.ALIGN_LEFT;
                        stfoto.Colspan = 2;
                        stfoto.Border = 0;
                        stfoto.Rowspan = 3;
                        stfoto.PaddingTop = 0.5f;
                        stfoto.PaddingBottom = 1f;
                        HeadT.AddCell(stfoto);

                        PdfPCell messag = new PdfPCell(new Phrase("The purpose of this report is to communicate student achievement in relationship"+Environment.NewLine+"to trimester goals as well as what is required for future progress toward them.", new Font(Font.FontFamily.HELVETICA, 11, Font.NORMAL, BaseColor.BLACK)));
                        messag.HorizontalAlignment = Element.ALIGN_LEFT;
                        messag.Border = 0;
                        messag.Colspan = 7;
                        messag.PaddingTop = 2;
                        messag.PaddingBottom = 5;
                        messag.Rowspan = 3;
                        HeadT.AddCell(messag);

                        PdfPCell grad = new PdfPCell(new Phrase("Grade: " + grade, new Font(Font.FontFamily.HELVETICA, 11, Font.BOLD, BaseColor.BLACK)));
                        grad.HorizontalAlignment = Element.ALIGN_LEFT;
                        grad.Border = 0;
                        grad.Colspan = 5;
                        HeadT.AddCell(grad);

                        PdfPCell stinu = new PdfPCell(new Phrase("StudentID: " + std, new Font(Font.FontFamily.HELVETICA, 5, Font.BOLD, BaseColor.WHITE)));
                        stinu.HorizontalAlignment = Element.ALIGN_LEFT;
                        stinu.Border = 0;
                        stinu.Colspan = 5;
                        stinu.PaddingBottom = 3;
                        HeadT.AddCell(stinu);
                       
                        PdfPCell HR = new PdfPCell(new Phrase("Homeroom: " + HT, new Font(Font.FontFamily.HELVETICA, 11, Font.BOLD, BaseColor.BLACK)));
                        HR.HorizontalAlignment = Element.ALIGN_LEFT;
                        HR.Border = 0;
                        HR.Colspan = 10;
                        HR.PaddingBottom = 5;
                        HeadT.AddCell(HR);
                        

                        PdfPCell bar2 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                        bar2.HorizontalAlignment = Element.ALIGN_LEFT;
                        bar2.Border = 0;
                        bar2.Colspan = 16;
                        bar2.BackgroundColor = new BaseColor(135, 0, 27);
                        HeadT.AddCell(bar2);

                        //Legend
                        PdfPTable legendTable = new PdfPTable(18);
                        legendTable.HorizontalAlignment = Element.ALIGN_CENTER;
                        legendTable.WidthPercentage = 100;

                        PdfPCell cel1 = new PdfPCell(new Phrase("STANDARDS PROFICIENCY KEY ", new Font(Font.FontFamily.HELVETICA, 9F, Font.BOLD, BaseColor.BLACK)));
                        cel1.HorizontalAlignment = Element.ALIGN_LEFT;
                        cel1.Colspan = 18;
                        cel1.BorderColor = BaseColor.LIGHT_GRAY;
                        cel1.PaddingTop = 5;
                        cel1.BorderWidthTop = 0;
                        cel1.BorderWidthBottom = 1;
                        cel1.BorderWidthRight = 0;
                        cel1.BorderWidthLeft = 0;
                        legendTable.AddCell(cel1);
                        PdfPCell cel2 = new PdfPCell(new Phrase("Code", new Font(Font.FontFamily.HELVETICA, 9F, Font.BOLD, BaseColor.BLACK)));
                        cel2.HorizontalAlignment = Element.ALIGN_LEFT;
                        cel2.BorderWidthLeft = 1;
                        cel2.BorderWidthBottom = 1;
                        cel2.BorderColor = BaseColor.LIGHT_GRAY;
                        cel2.BorderWidthRight = 0;
                        cel2.BorderWidthTop = 0;
                        legendTable.AddCell(cel2);
                        PdfPCell cel3 = new PdfPCell(new Phrase("Achievement Descriptors", new Font(Font.FontFamily.HELVETICA, 9F, Font.BOLD, BaseColor.BLACK)));
                        cel3.HorizontalAlignment = Element.ALIGN_LEFT;
                        cel3.BorderWidthBottom = 1;
                        cel3.BorderWidthRight = 0;
                        cel3.BorderWidthLeft = 0;
                        cel3.BorderWidthTop = 0;
                        cel3.BorderColor = BaseColor.LIGHT_GRAY;
                        cel3.Colspan = 5;
                        legendTable.AddCell(cel3);
                        PdfPCell cel4 = new PdfPCell(new Phrase("Behavioral Descriptors", new Font(Font.FontFamily.HELVETICA, 9F, Font.BOLD, BaseColor.BLACK)));
                        cel4.HorizontalAlignment = Element.ALIGN_LEFT;
                        cel4.Colspan = 12;
                        cel4.BorderWidthLeft = 0;
                        cel4.BorderWidthRight = 1;
                        cel4.BorderWidthTop = 0;
                        cel4.BorderColor = BaseColor.LIGHT_GRAY;
                        legendTable.AddCell(cel4);
                        PdfPCell cel1d = new PdfPCell(new Phrase("5"+Environment.NewLine+"4" + Environment.NewLine + "3" + Environment.NewLine + "2" + Environment.NewLine + "1" + Environment.NewLine + "--" + Environment.NewLine + "*", new Font(Font.FontFamily.HELVETICA, 8.0F, Font.NORMAL, BaseColor.BLACK)));
                        cel1d.HorizontalAlignment = Element.ALIGN_CENTER;
                        cel1d.BorderWidthLeft = 1;
                        cel1d.BorderWidthBottom = 1;
                        cel1d.BorderWidthRight = 0;
                        cel1d.BorderWidthTop = 0;
                        cel1d.BorderColor = BaseColor.LIGHT_GRAY;
                        legendTable.AddCell(cel1d);
                        PdfPCell cel2d = new PdfPCell(new Phrase("Meets Trimester Standard with Distinction" + Environment.NewLine + "Meets Trimester Standard" + Environment.NewLine + "Nearly Meets Trimester Standard" + Environment.NewLine + "Below Trimester Standard" + Environment.NewLine + "Far Below Trimester Standard" + Environment.NewLine + "Not Assessed This Trimester" + Environment.NewLine + "Based on Modified Expectations", new Font(Font.FontFamily.HELVETICA, 8F, Font.NORMAL, BaseColor.BLACK)));
                        cel2d.HorizontalAlignment = Element.ALIGN_LEFT;
                        cel2d.Colspan = 5;
                        cel2d.BorderWidthBottom = 1;
                        cel2d.BorderWidthTop = 0;
                        cel2d.BorderWidthLeft = 0;
                        cel2d.BorderWidthRight = 0;
                        cel2d.BorderColor = BaseColor.LIGHT_GRAY;
                        legendTable.AddCell(cel2d);
                        PdfPCell cel3d = new PdfPCell(new Phrase("The student takes understandings and learning beyong trimester benchmark consistantly." + Environment.NewLine + "The student knows and/or is able to do trimester benchmark consistently" + Environment.NewLine + "The student knows and/or is able to do trimester benchmark, but not consistently" + Environment.NewLine + "The student does not know and/or unable to do trimester benchmark, but shows beginning understandings." + Environment.NewLine + "The student does not know and/or is unable to do trimester benchmark." + Environment.NewLine + "The student was not assessed on this benchmark this trimester." + Environment.NewLine + "The student was assessed based on his/her individualized educational goals.", new Font(Font.FontFamily.HELVETICA, 8F, Font.NORMAL, BaseColor.BLACK)));
                        cel3d.HorizontalAlignment = Element.ALIGN_LEFT;
                        cel3d.Colspan= 12;
                        cel3d.BorderWidthLeft = 0;
                        cel3d.BorderWidthBottom = 1;
                        cel3d.BorderWidthRight = 1;
                        cel3d.BorderColor = BaseColor.LIGHT_GRAY;
                        legendTable.AddCell(cel3d);

                        /// GRADE DETAILS
                        PdfPTable GradeTable = new PdfPTable(20);
                        GradeTable.HorizontalAlignment = Element.ALIGN_CENTER;
                        GradeTable.WidthPercentage = 100;

                        var subt = "";
                        var TAPB = "";
                        var hed = "";
                        var hed2 = "";
                        var hrw = "";
                        for (int i = 0; i < stTable.Length-1; i++)
                        {
                            var hr = stTable[i].Split('|');
                            var idt = hr[11].Split('.')[2];
                            hed = hr[4];
                            if (hr[5]== "Handwriting")
                            {
                                hrw=hr[5]+"|"+hr[6]+"|" + hr[7]+"|" + hr[8]+"|" + hr[9];
                            }
                            if (hed == "" + grade + "LA")
                            {
                                TAPB = "ENGLISH LANGUAGE ARTS";
                            }
                            else if (hed == "" + grade + "MA" || hed == "" + grade + "Ma")
                            {
                                TAPB = "MATHEMATICS";
                            }
                            else if (hed == "" + grade + "SLA")
                            {
                                TAPB = "SPANISH LANGUAGE ARTS";
                            }
                            else if (hed == "" + grade + "HR")
                            {
                                TAPB = "CONDUCT";
                            }

                            else if (hed == "" + grade + "SC")
                            {
                                TAPB = "SCIENCE";
                            }
                            else if (hed == "" + grade + "SS")
                            {
                                TAPB = "SOCIAL STUDIES";
                            }
                            else if (hed == "" + grade + "PE")
                            {
                                TAPB = "PHYSICAL EDUCATION / HEALTH";
                            }
                            else if (hed == "PKSC,PKSS")
                            {
                                TAPB = "SOCIAL STUDIES";
                                hr[3] = HT;
                            }



                            if (hed != hed2) {
                               
                                PdfPCell spa2 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 5.0F, Font.NORMAL, BaseColor.BLACK)));
                                spa2.Border = 0;
                                spa2.Colspan = 20;
                                GradeTable.AddCell(spa2);

                                PdfPCell Course = new PdfPCell(new Phrase(TAPB, new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.WHITE)));
                                Course.HorizontalAlignment = Element.ALIGN_CENTER;
                                Course.BorderWidth = 1F;
                                Course.BackgroundColor = new BaseColor(135, 0, 27);
                                Course.Colspan = 6;
                                Course.BorderColor = BaseColor.GRAY;
                                GradeTable.AddCell(Course);
                                PdfPCell Teacher = new PdfPCell(new Phrase(hr[3], new Font(Font.FontFamily.HELVETICA, 11.0F, Font.ITALIC, BaseColor.BLACK)));
                                Teacher.HorizontalAlignment = Element.ALIGN_LEFT;
                                Teacher.BorderWidth = 1F;
                                Teacher.Colspan = 11;
                                Teacher.BorderColor = BaseColor.GRAY;
                                Teacher.Border = 0;
                                GradeTable.AddCell(Teacher);

                                
                                if (hr[4] == "" + grade + "HR")
                                {
                                    
                                    PdfPCell T1 = new PdfPCell(new Phrase("T1", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                                    T1.HorizontalAlignment = Element.ALIGN_LEFT;
                                    T1.BackgroundColor = new BaseColor(135, 0, 27);
                                    T1.BorderWidth = 1F;
                                    GradeTable.AddCell(T1);
                                    PdfPCell T2 = new PdfPCell(new Phrase("T2", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                                    T2.HorizontalAlignment = Element.ALIGN_LEFT;
                                    T2.BackgroundColor = new BaseColor(135, 0, 27);
                                    T2.BorderWidth = 1F;
                                    GradeTable.AddCell(T2);
                                    PdfPCell T3 = new PdfPCell(new Phrase("T3", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                                    T3.HorizontalAlignment = Element.ALIGN_LEFT;
                                    T3.BackgroundColor = new BaseColor(135, 0, 27);
                                    T3.BorderWidth = 1F;
                                    GradeTable.AddCell(T3);
                                }else
                                {
                                    PdfPCell spa = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                    spa.HorizontalAlignment = Element.ALIGN_LEFT;
                                    spa.BorderWidth = 1F;
                                    spa.Colspan = 3;
                                    spa.Border = 0;
                                    GradeTable.AddCell(spa);
                                }

                                if (idt == "LA" && hrw != "")
                                {
                                    PdfPCell subj1 = new PdfPCell(new Phrase(hrw.Split('|')[0], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                    subj1.HorizontalAlignment = Element.ALIGN_LEFT;
                                    subj1.BorderWidth = 1F;
                                    subj1.Colspan = 4;
                                    subj1.BorderColor = BaseColor.GRAY;
                                    GradeTable.AddCell(subj1);

                                    PdfPCell stnam1 = new PdfPCell(new Phrase(hrw.Split('|')[1], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                    stnam1.HorizontalAlignment = Element.ALIGN_LEFT;
                                    stnam1.BorderWidth = 1F;
                                    stnam1.Colspan = 13;
                                    stnam1.BorderColor = BaseColor.GRAY;
                                    GradeTable.AddCell(stnam1);

                                    
                                        PdfPCell vt12 = new PdfPCell(new Phrase(hrw.Split('|')[2], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        vt12.HorizontalAlignment = Element.ALIGN_CENTER;
                                        vt12.BorderWidth = 1F;
                                        vt12.BorderColor = BaseColor.GRAY;
                                        GradeTable.AddCell(vt12);
                                  
                                        PdfPCell vt22 = new PdfPCell(new Phrase(hrw.Split('|')[3], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        vt22.HorizontalAlignment = Element.ALIGN_CENTER;
                                        vt22.BorderWidth = 1F;
                                        vt22.BorderColor = BaseColor.GRAY;
                                        GradeTable.AddCell(vt22);

                                 
                                        PdfPCell vt31 = new PdfPCell(new Phrase(hrw.Split('|')[4], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        vt31.HorizontalAlignment = Element.ALIGN_CENTER;
                                        vt31.BorderWidth = 1F;
                                        vt31.BorderColor = BaseColor.GRAY;
                                        GradeTable.AddCell(vt31);

                                    hrw = "";
                                }
                                    
                                    hed2 = hr[4];

                                
                            }

                            if (hr[5]=="Tribes" || hr[5] == "TRIBES" || hr[5]== "TribesTLCÆ")
                            {
                                hr[5] = "Tribes® Agreements";
                        }else if(hr[5] == "Listening and Speaking.")
                            {
                                hr[5] = "Listening and Speaking";
                            }else if(hr[5]== "SOCIAL/WORK DEVELOPMENT (TRIBES®)" || hr[5] == "Social/Work Development (TRIBES)")
                            {
                                hr[5] = "Tribes® Agreements";
                            }

                            if (hr[5]!= "Handwriting") { 
                                
                                if (subt != hr[5] )
                            {

                                    PdfPCell subj = new PdfPCell(new Phrase(hr[5], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                subj.HorizontalAlignment = Element.ALIGN_LEFT;
                                subj.BorderWidth = 1F;
                                subj.Colspan = 4;
                                subj.BorderColor = BaseColor.GRAY;
                                GradeTable.AddCell(subj);
                                    subt = hr[5];
                            }
                            else
                            {

                                PdfPCell subj = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                subj.HorizontalAlignment = Element.ALIGN_LEFT;
                                subj.Border=0;
                                subj.Colspan = 4;
                                GradeTable.AddCell(subj);
                            }
                                 
                                    PdfPCell stnam = new PdfPCell(new Phrase(hr[6], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                    stnam.HorizontalAlignment = Element.ALIGN_LEFT;
                                    stnam.BorderWidth = 1F;
                                    stnam.Colspan = 13;
                            stnam.BorderColor = BaseColor.GRAY;
                            GradeTable.AddCell(stnam);
                               
                                    PdfPCell vt1 = new PdfPCell(new Phrase(hr[7], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                    vt1.HorizontalAlignment = Element.ALIGN_CENTER;
                                    vt1.BorderWidth = 1F;
                            vt1.BorderColor = BaseColor.GRAY;
                            GradeTable.AddCell(vt1);
                                

                             
                                    PdfPCell vt2 = new PdfPCell(new Phrase(hr[8], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                            vt2.HorizontalAlignment = Element.ALIGN_CENTER;
                            vt2.BorderWidth = 1F;
                            vt2.BorderColor = BaseColor.GRAY;
                            GradeTable.AddCell(vt2);
                                
                             
                                    PdfPCell vt3 = new PdfPCell(new Phrase(hr[9], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                            vt3.HorizontalAlignment = Element.ALIGN_CENTER;
                            vt3.BorderWidth = 1F;
                            vt3.BorderColor = BaseColor.GRAY;
                            GradeTable.AddCell(vt3);
                                
                                
                            }
                        }
                        PdfPTable espTable = new PdfPTable(18);
                        espTable.HorizontalAlignment = Element.ALIGN_CENTER;
                        espTable.WidthPercentage = 100;
                        if (T1ESP != "") { 
                        var espT = T1ESP.Split('^');

                        PdfPCell ep1 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                        ep1.HorizontalAlignment = Element.ALIGN_LEFT;
                        ep1.Colspan = 18;
                        ep1.Border = 0;
                        espTable.AddCell(ep1);
                        PdfPCell espe = new PdfPCell(new Phrase("Fine Arts & Technology", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                        espe.HorizontalAlignment = Element.ALIGN_LEFT;
                        espe.BackgroundColor = new BaseColor(135, 0, 27);
                        espe.BorderWidth = 1F;
                        espe.Colspan = 12;
                        espTable.AddCell(espe);
                        PdfPCell ST1 = new PdfPCell(new Phrase("T1", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                        ST1.HorizontalAlignment = Element.ALIGN_CENTER;
                        ST1.BackgroundColor = new BaseColor(135, 0, 27);
                        ST1.BorderWidth = 1F;
                        ST1.Colspan = 2;
                        espTable.AddCell(ST1);
                        PdfPCell ST2 = new PdfPCell(new Phrase("T2", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                        ST2.HorizontalAlignment = Element.ALIGN_CENTER;
                        ST2.BackgroundColor = new BaseColor(135, 0, 27);
                        ST2.BorderWidth = 1F;
                        ST2.Colspan =2;
                        espTable.AddCell(ST2);
                        PdfPCell ST3 = new PdfPCell(new Phrase("T3", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                        ST3.HorizontalAlignment = Element.ALIGN_CENTER;
                        ST3.BackgroundColor = new BaseColor(135, 0, 27);
                        ST3.BorderWidth = 1F;
                        ST3.Colspan = 2;
                        espTable.AddCell(ST3);
                        PdfPCell SP = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                        SP.HorizontalAlignment = Element.ALIGN_LEFT;
                        SP.BackgroundColor = new BaseColor(135, 0, 27);
                        SP.BorderWidth = 1F;
                        SP.Colspan = 12;
                        espTable.AddCell(SP);
                        PdfPCell U1 = new PdfPCell(new Phrase("U", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                        U1.HorizontalAlignment = Element.ALIGN_CENTER;
                        U1.BackgroundColor = new BaseColor(135, 0, 27);
                        U1.BorderWidth = 1F;
                        espTable.AddCell(U1);
                        PdfPCell SPT1 = new PdfPCell(new Phrase("T", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                        SPT1.HorizontalAlignment = Element.ALIGN_CENTER;
                        SPT1.BackgroundColor = new BaseColor(135, 0, 27);
                        SPT1.BorderWidth = 1F;
                        espTable.AddCell(SPT1);
                        PdfPCell U2 = new PdfPCell(new Phrase("U", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                        U2.HorizontalAlignment = Element.ALIGN_CENTER;
                        U2.BackgroundColor = new BaseColor(135, 0, 27);
                        U2.BorderWidth = 1F;
                        espTable.AddCell(U2);
                        PdfPCell SPT2 = new PdfPCell(new Phrase("T", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                        SPT2.HorizontalAlignment = Element.ALIGN_CENTER;
                        SPT2.BackgroundColor = new BaseColor(135, 0, 27);
                        SPT2.BorderWidth = 1F;
                        espTable.AddCell(SPT2);
                        PdfPCell U3 = new PdfPCell(new Phrase("U", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                        U3.HorizontalAlignment = Element.ALIGN_CENTER;
                        U3.BackgroundColor = new BaseColor(135, 0, 27);
                        U3.BorderWidth = 1F;
                        espTable.AddCell(U3);
                        PdfPCell SPT3 = new PdfPCell(new Phrase("T", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                        SPT3.HorizontalAlignment = Element.ALIGN_CENTER;
                        SPT3.BackgroundColor = new BaseColor(135, 0, 27);
                        SPT3.BorderWidth = 1F;
                        espTable.AddCell(SPT3);
                            var cour1 = "";
                            for (int a = 0; a < espT.Length-1; a++)
                        {
                            var esVal = espT[a].Split('|');
                            var cour = "";
                            if (esVal[0] == "" + grade + "TECH")
                            {
                                cour = "TECHNOLOGY";
                            }else if(esVal[0] == "" + grade + "ART" || esVal[0] == "" + grade + "Art")
                            {
                                cour = "ART";
                            }
                           
                            else if (esVal[0] == "" + grade + "MUS" || esVal[0] == "" + grade + "Mus")
                            {
                                cour = "MUSIC";
                            }
                                
                            if (cour != cour1)
                                {
                                   
                                        PdfPCell cou = new PdfPCell(new Phrase(cour, new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        cou.HorizontalAlignment = Element.ALIGN_LEFT;
                                        cou.BorderWidth = 1F;
                                        cou.Colspan = 6;
                                        cou.BorderColor = BaseColor.GRAY;
                                        espTable.AddCell(cou);
                                        PdfPCell tea = new PdfPCell(new Phrase(esVal[1], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.ITALIC, BaseColor.BLACK)));
                                        tea.HorizontalAlignment = Element.ALIGN_LEFT;
                                        tea.BorderWidth = 1F;
                                        tea.Colspan = 6;
                                        tea.BorderColor = BaseColor.GRAY;
                                        espTable.AddCell(tea);
                                        PdfPCell SPut1;
                                        if (esVal[2] != "")
                                        {
                                            SPut1 = new PdfPCell(new Phrase(esVal[2].Split('/')[1], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        }
                                        else
                                        {
                                            SPut1 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        }
                                        SPut1.HorizontalAlignment = Element.ALIGN_CENTER;
                                        SPut1.BorderWidth = 1F;
                                        SPut1.BorderColor = BaseColor.GRAY;
                                        espTable.AddCell(SPut1);
                                        PdfPCell SPut2;
                                        if (esVal[2] != "")
                                        {
                                            SPut2 = new PdfPCell(new Phrase(esVal[2].Split('/')[0], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        }
                                        else
                                        {
                                            SPut2 = new PdfPCell(new Phrase("", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        }
                                        SPut2.HorizontalAlignment = Element.ALIGN_CENTER;
                                        SPut2.BorderWidth = 1F;
                                        SPut2.BorderColor = BaseColor.GRAY;
                                        espTable.AddCell(SPut2);
                                        PdfPCell SPut3;
                                        if (esVal[3] != "")
                                        {
                                            SPut3 = new PdfPCell(new Phrase(esVal[3].Split('/')[1], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        }
                                        else
                                        {
                                            SPut3 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        }
                                        SPut3.HorizontalAlignment = Element.ALIGN_CENTER;
                                        SPut3.BorderWidth = 1F;
                                        SPut3.BorderColor = BaseColor.GRAY;
                                        espTable.AddCell(SPut3);
                                        PdfPCell SPut4;
                                        if (esVal[3] != "")
                                        {
                                            SPut4 = new PdfPCell(new Phrase(esVal[3].Split('/')[0], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        }
                                        else
                                        {
                                            SPut4 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        }

                                        SPut4.HorizontalAlignment = Element.ALIGN_CENTER;
                                        SPut4.BorderWidth = 1F;
                                        SPut4.BorderColor = BaseColor.GRAY;
                                        espTable.AddCell(SPut4);
                                        PdfPCell SPut5;
                                        if (esVal[4] != "")
                                        {
                                            SPut5 = new PdfPCell(new Phrase(esVal[4].Split('/')[1], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        }
                                        else
                                        {
                                            SPut5 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        }
                                        SPut5.HorizontalAlignment = Element.ALIGN_CENTER;
                                        SPut5.BorderWidth = 1F;
                                        SPut5.BorderColor = BaseColor.GRAY;
                                        espTable.AddCell(SPut5);
                                        PdfPCell SPut6;
                                        if (esVal[4] != "")
                                        {
                                            SPut6 = new PdfPCell(new Phrase(esVal[4].Split('/')[0], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        }
                                        else
                                        {
                                            SPut6 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                        }
                                        SPut6.HorizontalAlignment = Element.ALIGN_CENTER;
                                        SPut6.BorderWidth = 1F;
                                        SPut6.BorderColor = BaseColor.GRAY;
                                        espTable.AddCell(SPut6);
                                        cour1 = cour;
                                   
                                }
                            
                        }
                        PdfPCell fo = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.WHITE)));
                        fo.HorizontalAlignment = Element.ALIGN_LEFT;
                        fo.Colspan = 18;
                        fo.Border = 0;
                        espTable.AddCell(fo);
                       
                        PdfPCell ATTE = new PdfPCell(new Phrase("ATTENDANCE", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                        ATTE.HorizontalAlignment = Element.ALIGN_LEFT;
                        ATTE.Colspan = 6;
                        ATTE.Border = 1;
                        ATTE.BorderColor = BaseColor.GRAY;
                        ATTE.BackgroundColor = new BaseColor(135, 0, 27);
                        espTable.AddCell(ATTE);
                        PdfPCell A1 = new PdfPCell(new Phrase("T1", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                        A1.HorizontalAlignment = Element.ALIGN_CENTER;
                        A1.BackgroundColor = new BaseColor(135, 0, 27);
                        A1.BorderWidth = 1F;
                        A1.Border = 1;
                        A1.BorderColor = BaseColor.GRAY;
                        espTable.AddCell(A1);
                        PdfPCell A2 = new PdfPCell(new Phrase("T2", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                        A2.HorizontalAlignment = Element.ALIGN_CENTER;
                        A2.BackgroundColor = new BaseColor(135, 0, 27);
                        A2.BorderWidth = 1F;
                        A2.Border = 1;
                        A2.BorderColor = BaseColor.GRAY;
                        espTable.AddCell(A2);
                        PdfPCell A3 = new PdfPCell(new Phrase("T3", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                        A3.HorizontalAlignment = Element.ALIGN_CENTER;
                        A3.BackgroundColor = new BaseColor(135, 0, 27);
                        A3.BorderWidth = 1F;
                        A3.Border = 1;
                        A3.BorderColor = BaseColor.GRAY;
                        espTable.AddCell(A3);
                        PdfPCell spa5 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.WHITE)));
                        spa5.Colspan = 3;
                        spa5.Rowspan = 3;
                        spa5.Border = 0;
                        espTable.AddCell(spa5);
                        PdfPCell STK = new PdfPCell(new Phrase("STANDARD KEY", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                        STK.HorizontalAlignment = Element.ALIGN_LEFT;
                        STK.Colspan = 6;
                        STK.Border = 1;
                        STK.BorderWidth = 1F;
                        STK.BorderColor = BaseColor.GRAY;
                        STK.BackgroundColor = new BaseColor(135, 0, 27);
                        espTable.AddCell(STK);
                        PdfPCell ABSEN = new PdfPCell(new Phrase("Days Absent", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                        ABSEN.HorizontalAlignment = Element.ALIGN_LEFT;
                        ABSEN.Colspan = 6;
                        ABSEN.BorderColor = BaseColor.GRAY;
                        espTable.AddCell(ABSEN);
                            if (ABTAR != "")
                            {
                                if (ABTAR.Split('|')[0] != "")
                                {
                                    PdfPCell abt = new PdfPCell(new Phrase(ABTAR.Split('|')[0].Split(',')[0], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                    abt.HorizontalAlignment = Element.ALIGN_CENTER;
                                    abt.BorderColor = BaseColor.GRAY;
                                    espTable.AddCell(abt);
                                }
                                else
                                {
                                    PdfPCell abt = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                    abt.HorizontalAlignment = Element.ALIGN_CENTER;
                                    abt.BorderColor = BaseColor.GRAY;
                                    espTable.AddCell(abt);
                                }

                                if (ABTAR.Split('|')[1] != "")
                                {

                                    PdfPCell abt1 = new PdfPCell(new Phrase(ABTAR.Split('|')[1].Split(',')[0], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                    abt1.HorizontalAlignment = Element.ALIGN_CENTER;
                                    abt1.BorderColor = BaseColor.GRAY;
                                    espTable.AddCell(abt1);
                                }
                                else
                                {
                                    PdfPCell abt1 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.WHITE)));
                                    abt1.HorizontalAlignment = Element.ALIGN_CENTER;
                                    abt1.BorderColor = BaseColor.GRAY;
                                    espTable.AddCell(abt1);
                                }
                                if (ABTAR.Split('|')[2] != "")
                                {
                                    PdfPCell abt2 = new PdfPCell(new Phrase(ABTAR.Split('|')[2].Split(',')[0], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                    abt2.HorizontalAlignment = Element.ALIGN_CENTER;
                                    abt2.BorderColor = BaseColor.GRAY;
                                    espTable.AddCell(abt2);
                                }
                                else
                                {
                                    PdfPCell abt2 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.WHITE)));
                                    abt2.HorizontalAlignment = Element.ALIGN_CENTER;
                                    abt2.BorderColor = BaseColor.GRAY;
                                    espTable.AddCell(abt2);
                                }
                            }
                            else
                            {
                                PdfPCell tard1 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.WHITE)));
                                tard1.HorizontalAlignment = Element.ALIGN_CENTER;
                                tard1.BorderColor = BaseColor.GRAY;
                                espTable.AddCell(tard1);
                                PdfPCell tard2 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.WHITE)));
                                tard2.HorizontalAlignment = Element.ALIGN_CENTER;
                                tard2.BorderColor = BaseColor.GRAY;
                                espTable.AddCell(tard2);
                                PdfPCell tard3 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.WHITE)));
                                tard3.HorizontalAlignment = Element.ALIGN_CENTER;
                                tard3.BorderColor = BaseColor.GRAY;
                                espTable.AddCell(tard3);
                            }
                            PdfPCell leg = new PdfPCell(new Phrase("U = Understands concepts and uses skills." + Environment.NewLine + "T = Follows Tribes® Agreements.", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                            leg.HorizontalAlignment = Element.ALIGN_LEFT;
                            leg.Colspan = 6;
                            leg.Rowspan = 2;
                            leg.Border = 1;
                            leg.BorderColor = BaseColor.GRAY;
                            espTable.AddCell(leg);

                            PdfPCell TARD = new PdfPCell(new Phrase("Days Tardy", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                            TARD.HorizontalAlignment = Element.ALIGN_LEFT;
                            TARD.BorderColor = BaseColor.GRAY;
                            TARD.Colspan = 6;
                            espTable.AddCell(TARD);
                            if (ABTAR != "")
                            {
                                if (ABTAR.Split('|')[0] != "")
                                {
                                    PdfPCell tard1 = new PdfPCell(new Phrase(ABTAR.Split('|')[0].Split(',')[1], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                    tard1.HorizontalAlignment = Element.ALIGN_CENTER;
                                    tard1.BorderColor = BaseColor.GRAY;
                                    espTable.AddCell(tard1);
                                }
                                else
                                {
                                    PdfPCell tard1 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                    tard1.HorizontalAlignment = Element.ALIGN_CENTER;
                                    tard1.BorderColor = BaseColor.GRAY;
                                    espTable.AddCell(tard1);
                                }
                                if (ABTAR.Split('|')[1] != "")
                                {
                                    PdfPCell tard2 = new PdfPCell(new Phrase(ABTAR.Split('|')[1].Split(',')[1], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                    tard2.HorizontalAlignment = Element.ALIGN_CENTER;
                                    tard2.BorderColor = BaseColor.GRAY;
                                    espTable.AddCell(tard2);
                                }
                                else
                                {
                                    PdfPCell tard2 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.WHITE)));
                                    tard2.HorizontalAlignment = Element.ALIGN_CENTER;
                                    tard2.BorderColor = BaseColor.GRAY;
                                    espTable.AddCell(tard2);
                                }
                                if (ABTAR.Split('|')[2] != "")
                                {
                                    PdfPCell tard4 = new PdfPCell(new Phrase(ABTAR.Split('|')[2].Split(',')[1], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                    tard4.HorizontalAlignment = Element.ALIGN_CENTER;
                                    tard4.BorderColor = BaseColor.GRAY;
                                    espTable.AddCell(tard4);
                                }
                                else
                                {
                                    PdfPCell tard4 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.WHITE)));
                                    tard4.HorizontalAlignment = Element.ALIGN_CENTER;
                                    tard4.BorderColor = BaseColor.GRAY;
                                    espTable.AddCell(tard4);
                                }

                            }
                            else
                            {
                                PdfPCell tard1 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.WHITE)));
                                tard1.HorizontalAlignment = Element.ALIGN_CENTER;
                                tard1.BorderColor = BaseColor.GRAY;
                                espTable.AddCell(tard1);
                                PdfPCell tard2 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.WHITE)));
                                tard2.HorizontalAlignment = Element.ALIGN_CENTER;
                                tard2.BorderColor = BaseColor.GRAY;
                                espTable.AddCell(tard2);
                                PdfPCell tard3 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.WHITE)));
                                tard3.HorizontalAlignment = Element.ALIGN_CENTER;
                                tard3.BorderColor = BaseColor.GRAY;
                                espTable.AddCell(tard3);
                            }
                        }
                        else
                        {
                            PdfPCell ep21 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                            ep21.HorizontalAlignment = Element.ALIGN_LEFT;
                            ep21.Colspan = 18;
                            ep21.Border = 0;
                            espTable.AddCell(ep21);
                            PdfPCell ATTE = new PdfPCell(new Phrase("ATTENDANCE", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                            ATTE.HorizontalAlignment = Element.ALIGN_LEFT;
                            ATTE.Colspan = 15;
                            ATTE.Border = 1;
                            ATTE.BorderColor = BaseColor.GRAY;
                            ATTE.BackgroundColor = new BaseColor(135, 0, 27);
                            espTable.AddCell(ATTE);
                            PdfPCell A1 = new PdfPCell(new Phrase("T1", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                            A1.HorizontalAlignment = Element.ALIGN_CENTER;
                            A1.BackgroundColor = new BaseColor(135, 0, 27);
                            A1.BorderWidth = 1F;
                            A1.Border = 1;
                            A1.BorderColor = BaseColor.GRAY;
                            espTable.AddCell(A1);
                            PdfPCell A2 = new PdfPCell(new Phrase("T2", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                            A2.HorizontalAlignment = Element.ALIGN_CENTER;
                            A2.BackgroundColor = new BaseColor(135, 0, 27);
                            A2.BorderWidth = 1F;
                            A2.Border = 1;
                            A2.BorderColor = BaseColor.GRAY;
                            espTable.AddCell(A2);
                            PdfPCell A3 = new PdfPCell(new Phrase("T3", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                            A3.HorizontalAlignment = Element.ALIGN_CENTER;
                            A3.BackgroundColor = new BaseColor(135, 0, 27);
                            A3.BorderWidth = 1F;
                            A3.Border = 1;
                            A3.BorderColor = BaseColor.GRAY;
                            espTable.AddCell(A3);
                            PdfPCell ABSEN = new PdfPCell(new Phrase("Days Absent", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                            ABSEN.HorizontalAlignment = Element.ALIGN_LEFT;
                            ABSEN.Colspan = 15;
                            ABSEN.BorderColor = BaseColor.GRAY;
                            espTable.AddCell(ABSEN);
                            if (ABTAR != "")

                            {
                                if (ABTAR.Split('|')[0] != "")
                                {

                                    PdfPCell abt = new PdfPCell(new Phrase(ABTAR.Split('|')[0].Split(',')[0], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                    abt.HorizontalAlignment = Element.ALIGN_CENTER;
                                    abt.BorderColor = BaseColor.GRAY;
                                    espTable.AddCell(abt);
                                }
                                else
                                {
                                    PdfPCell abt = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                    abt.HorizontalAlignment = Element.ALIGN_CENTER;
                                    abt.BorderColor = BaseColor.GRAY;
                                    espTable.AddCell(abt);
                                }
                                if (ABTAR.Split('|')[1] != "")
                                {
                                    PdfPCell abt1 = new PdfPCell(new Phrase(ABTAR.Split('|')[1].Split(',')[0], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                    abt1.HorizontalAlignment = Element.ALIGN_CENTER;
                                    abt1.BorderColor = BaseColor.GRAY;
                                    espTable.AddCell(abt1);
                                }
                                else
                                {
                                    PdfPCell abt1 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.WHITE)));
                                    abt1.HorizontalAlignment = Element.ALIGN_CENTER;
                                    abt1.BorderColor = BaseColor.GRAY;
                                    espTable.AddCell(abt1);
                                }
                                if (ABTAR.Split('|')[2] != "")
                                {
                                    PdfPCell abt2 = new PdfPCell(new Phrase(ABTAR.Split('|')[2].Split(',')[0], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                    abt2.HorizontalAlignment = Element.ALIGN_CENTER;
                                    abt2.BorderColor = BaseColor.GRAY;
                                    espTable.AddCell(abt2);
                                }
                                else
                                {
                                    PdfPCell abt2 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.WHITE)));
                                    abt2.HorizontalAlignment = Element.ALIGN_CENTER;
                                    abt2.BorderColor = BaseColor.GRAY;
                                    espTable.AddCell(abt2);
                                }

                            }
                            PdfPCell TARD = new PdfPCell(new Phrase("Days Tardy", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                            TARD.HorizontalAlignment = Element.ALIGN_LEFT;
                            TARD.BorderColor = BaseColor.GRAY;
                            TARD.Colspan = 15;
                            espTable.AddCell(TARD);
                            if (ABTAR != "")
                            {
                                if (ABTAR.Split('|')[0] != "")
                                {
                                    PdfPCell tard1 = new PdfPCell(new Phrase(ABTAR.Split('|')[0].Split(',')[1], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                    tard1.HorizontalAlignment = Element.ALIGN_CENTER;
                                    tard1.BorderColor = BaseColor.GRAY;
                                    espTable.AddCell(tard1);
                                }
                                else
                                {
                                    PdfPCell tard1 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                    tard1.HorizontalAlignment = Element.ALIGN_CENTER;
                                    tard1.BorderColor = BaseColor.GRAY;
                                    espTable.AddCell(tard1);
                                }
                                if (ABTAR.Split('|')[1] != "")
                                {
                                    PdfPCell tard2 = new PdfPCell(new Phrase(ABTAR.Split('|')[1].Split(',')[1], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                    tard2.HorizontalAlignment = Element.ALIGN_CENTER;
                                    tard2.BorderColor = BaseColor.GRAY;
                                    espTable.AddCell(tard2);
                                }
                                else
                                {
                                    PdfPCell tard2 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.WHITE)));
                                    tard2.HorizontalAlignment = Element.ALIGN_CENTER;
                                    tard2.BorderColor = BaseColor.GRAY;
                                    espTable.AddCell(tard2);
                                }
                                if (ABTAR.Split('|')[2] != "")
                                {
                                    PdfPCell tard4 = new PdfPCell(new Phrase(ABTAR.Split('|')[2].Split(',')[1], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                    tard4.HorizontalAlignment = Element.ALIGN_CENTER;
                                    tard4.BorderColor = BaseColor.GRAY;
                                    espTable.AddCell(tard4);
                                }
                                else
                                {
                                    PdfPCell tard4 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.WHITE)));
                                    tard4.HorizontalAlignment = Element.ALIGN_CENTER;
                                    tard4.BorderColor = BaseColor.GRAY;
                                    espTable.AddCell(tard4);
                                }
                            }
                        }
                        PdfPCell ap6 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.WHITE)));
                        ap6.HorizontalAlignment = Element.ALIGN_CENTER;
                        ap6.Colspan = 18;
                        ap6.Border = 0;
                        espTable.AddCell(ap6);
                        PdfPCell comm1 = new PdfPCell(new Phrase("T1 Comment ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.BLACK)));
                        comm1.HorizontalAlignment = Element.ALIGN_LEFT;
                        comm1.Colspan = 18;
                        comm1.Border = 0;
                        comm1.BorderWidthBottom = 1;
                        comm1.BorderColorBottom = BaseColor.LIGHT_GRAY;
                        espTable.AddCell(comm1);

                        PdfPCell comms = new PdfPCell(new Phrase(comm, new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                        comms.HorizontalAlignment = Element.ALIGN_LEFT;
                        comms.Colspan = 18;
                        comms.Border = 0;
                        espTable.AddCell(comms);



                        documento.Add(HeadT);
                        documento.Add(legendTable); 
                       documento.Add(GradeTable);
                        documento.Add(espTable);

                        //Process prc = new System.Diagnostics.Process();
                        //prc.StartInfo.FileName = fileName;
                        //prc.Start();
                    }
                    else
                    {
                        con.Close();
                        fname = "";
                    }
                    con.Close();

                }

                documento.Close();
                con.Dispose();
            }
            catch (Exception ex)
            {
                throw;
            }
            return fname;
        }


        [WebMethod]
        public static string EXP_REPORTQ1(string stnum)
        {
            string sql = string.Empty;



            string fname = string.Empty;
            string fileName = string.Empty;
            OracleConnection con = new OracleConnection();
            con.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings["conn"].ConnectionString;
            Document documento = new Document(PageSize.LETTER, 10, 10, 5, 5);
            try
            {

                if (stnum.IndexOf(';') > -1)
                {
                    var stnumb = stnum.Split(';');
                    fname = "MS_ProgressReport_" + DateTime.Now.DayOfYear + DateTime.Now.Month + DateTime.Now.Day + DateTime.Now.Millisecond + ".pdf";
                    fileName = HttpContext.Current.Server.MapPath("~/RepoFiles/" + fname);
                    PdfWriter.GetInstance(documento, new FileStream(fileName, FileMode.Create));
                    documento.Open();

                    for (int a = 0; a < stnumb.Length; a++)
                    {
                        string Q1DATA = string.Empty;
                        string Q1AD = string.Empty;
                        string EXPDATA = string.Empty;
                        if (con.State == ConnectionState.Closed)
                        {
                            con.Open();
                        }

                        //RESP.

                        sql = "CREATE OR REPLACE VIEW MS_PRO_REP_RES";
                        sql += " AS SELECT S.ID,SEC.ID AS SECTID,S.STUDENT_NUMBER,s.Last_Name,s.First_name,C.COURSE_NAME,SG.STANDARDGRADE AS RES FROM standardgradesection sg";
                        sql += " LEFT JOIN STANDARD D ON sg.standardid = D.standardid";
                        sql += " LEFT JOIN STANDARD Di ON sg.standardid = Di.standardid";
                        sql += " LEFT JOIN SECTIONS SEC ON SG.SECTIONSDCID = SEC.dcid";
                        sql += " LEFT JOIN COURSES C ON SEC.COURSE_NUMBER = C.COURSE_NUMBER";
                        sql += " LEFT JOIN TEACHERS T ON SEC.TEACHER = T.id";
                        sql += " LEFT JOIN STUDENTS s ON sg.studentsdcid = s.dcid";
                        sql += " WHERE D.identifier LIKE '%RES%' AND C.COURSE_NAME NOT LIKE '%Explora%' AND C.COURSE_NAME NOT LIKE '%Space%' AND C.COURSE_NAME NOT LIKE '%Boot%' AND sg.yearid = 28";
                        sql += " AND sg.storecode IN ('Q1') AND SG.SCHOOLSDCID = 5 AND S.STUDENT_NUMBER =" + stnumb[a] + "";


                        OracleCommand cmdV1 = new OracleCommand(sql, con);
                        cmdV1.ExecuteNonQuery();

                        //CONDUCT.
                        sql = "CREATE OR REPLACE VIEW MS_PRO_REP_CON";
                        sql += " AS SELECT S.ID,SEC.ID AS SECID,S.STUDENT_NUMBER,s.Last_Name,s.First_name,C.COURSE_NAME,SG.STANDARDGRADE AS COND FROM standardgradesection sg";
                        sql += " LEFT JOIN STANDARD D ON sg.standardid = D.standardid";
                        sql += " LEFT JOIN SECTIONS SEC ON SG.SECTIONSDCID = SEC.dcid";
                        sql += " LEFT JOIN COURSES C ON SEC.COURSE_NUMBER = C.COURSE_NUMBER";
                        sql += " LEFT JOIN TEACHERS T ON SEC.TEACHER = T.id";
                        sql += " LEFT JOIN STUDENTS s ON sg.studentsdcid = s.dcid";
                        sql += " WHERE D.identifier LIKE '%CON%' AND C.COURSE_NAME NOT LIKE '%Explora%' AND C.COURSE_NAME NOT LIKE '%Advisory%' AND C.COURSE_NAME NOT LIKE '%Space%' AND C.COURSE_NAME NOT LIKE '%Boot%' AND sg.yearid = 28";
                        sql += " AND sg.storecode IN ('Q1') AND SG.SCHOOLSDCID = 5 AND S.STUDENT_NUMBER = " + stnumb[a] + "";


                        OracleCommand cmdV2 = new OracleCommand(sql, con);
                        cmdV2.ExecuteNonQuery();


                        //Advisory Teacher
                        sql = "SELECT DISTINCT C.COURSE_NAME,T.FIRST_NAME||' '||T.LAST_NAME AS TEACHER,R.RES,S.STUDENT_NUMBER,S.FIRST_NAME||' '||S.LAST_NAME AS STUDENT,S.GRADE_LEVEL FROM CC CO";
                        sql += " LEFT JOIN STUDENTS S ON CO.STUDENTID = S.ID";
                        sql += " LEFT JOIN COURSES C ON CO.COURSE_NUMBER = C.COURSE_NUMBER";
                        sql += " LEFT JOIN TEACHERS T ON CO.TEACHERID = T.ID";
                        sql += " LEFT JOIN MS_PRO_REP_RES R ON CO.STUDENTID = R.ID AND CO.SECTIONID = R.SECTID";
                        sql += " WHERE CO.TERMID IN(2800, 2801, 2802)  AND C.COURSE_NAME LIKE '%Advisory%' AND S.STUDENT_NUMBER =" + stnumb[a] + "";

                        OracleCommand cmd = new OracleCommand(sql, con);
                        OracleDataReader odr = cmd.ExecuteReader();
                        while (odr.Read())
                        {
                            Q1AD += odr["COURSE_NAME"].ToString() + '|';
                            Q1AD += odr["TEACHER"].ToString() + '|';
                            Q1AD += odr["RES"].ToString() + '|';
                            Q1AD += odr["STUDENT_NUMBER"].ToString() + '|';
                            Q1AD += odr["STUDENT"].ToString() + '|';
                            Q1AD += odr["GRADE_LEVEL"].ToString() + '|';

                        }


                        //GRADES VALUES
                        sql = "WITH main_query AS ( SELECT C.COURSE_NAME,T.FIRST_NAME||' '||T.LAST_NAME AS TEACHER, PG.FINALGRADENAME,PG.GRADE,R.RES,CN.COND,";
                        sql += " TO_CHAR(PG.COMMENT_VALUE) AS COMMENTS,CO.CURRENTABSENCES,CO.CURRENTTARDIES FROM CC CO";
                        sql += " LEFT JOIN STUDENTS S ON CO.STUDENTID=S.ID";
                        sql += " LEFT JOIN COURSES C ON CO.COURSE_NUMBER=C.COURSE_NUMBER";
                        sql += " LEFT JOIN TEACHERS T ON CO.TEACHERID=T.ID";
                        sql += " LEFT JOIN PGFINALGRADES PG ON CO.STUDENTID=PG.STUDENTID AND CO.SECTIONID=PG.SECTIONID AND FINALGRADENAME='Q1'";
                        sql += " LEFT JOIN MS_PRO_REP_RES R ON CO.STUDENTID = R.ID AND CO.SECTIONID = R.SECTID";
                        sql += " LEFT JOIN MS_PRO_REP_CON CN ON CO.STUDENTID = CN.ID AND CO.SECTIONID = CN.SECID";
                        sql += " WHERE CO.TERMID IN(2800,2801,2802) AND PG.FINALGRADENAME IN ('Q1') AND C.COURSE_NAME NOT LIKE '%Explora%' AND C.COURSE_NAME NOT LIKE '%Space%' AND C.COURSE_NAME NOT LIKE '%Advisory%' AND C.COURSE_NAME NOT LIKE '%Boot%' AND S.STUDENT_NUMBER=" + stnumb[a] + "";
                        sql += " )";
                        sql += " SELECT DISTINCT COURSE_NAME,TEACHER";
                        sql += " ,(SELECT  y.GRADE FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.FINALGRADENAME='Q1') SKILL";
                        sql += " ,(SELECT  y.RES FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.FINALGRADENAME='Q1') RESP";
                        sql += " ,(SELECT  y.COND FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.FINALGRADENAME='Q1') CONDT";
                        sql += " ,(SELECT  y.CURRENTABSENCES FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.FINALGRADENAME='Q1') ABS1";
                        sql += " ,(SELECT  y.CURRENTTARDIES FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.FINALGRADENAME = 'Q1') TARDI";
                        sql += " ,(SELECT  y.COMMENTS FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.FINALGRADENAME = 'Q1') Comments";
                        sql += "  FROM	main_query M";
                        sql += " ORDER BY COURSE_NAME";

                        OracleCommand cmd1 = new OracleCommand(sql, con);
                        OracleDataReader odr1 = cmd1.ExecuteReader();
                        while (odr1.Read())
                        {
                            Q1DATA += odr1["COURSE_NAME"].ToString() + '|';
                            Q1DATA += odr1["TEACHER"].ToString() + '|';
                            Q1DATA += odr1["SKILL"].ToString() + '|';
                            Q1DATA += odr1["RESP"].ToString() + '|';
                            Q1DATA += odr1["CONDT"].ToString() + '|';
                            Q1DATA += odr1["ABS1"].ToString() + '|';
                            Q1DATA += odr1["TARDI"].ToString() + '|';
                            Q1DATA += odr1["Comments"].ToString() + '^';

                        }

                        //EXPLORATORY RES.
                        sql = "CREATE OR REPLACE VIEW MS_PRO_EXP_RES";
                        sql += " AS SELECT S.ID AS STID,SEC.ID AS SECID,S.STUDENT_NUMBER,s.Last_Name,s.First_name,C.COURSE_NAME,SG.STANDARDGRADE AS RES FROM standardgradesection sg";
                        sql += " LEFT JOIN STANDARD D ON sg.standardid = D.standardid";
                        sql += " LEFT JOIN SECTIONS SEC ON SG.SECTIONSDCID = SEC.dcid";
                        sql += " LEFT JOIN COURSES C ON SEC.COURSE_NUMBER = C.COURSE_NUMBER";
                        sql += " LEFT JOIN TEACHERS T ON SEC.TEACHER = T.id";
                        sql += " LEFT JOIN STUDENTS s ON sg.studentsdcid = s.dcid";
                        sql += " WHERE D.identifier LIKE '%RES%' AND sg.yearid = 28 AND";
                        sql += " (C.COURSE_NAME LIKE '%Explora%' OR C.COURSE_NAME LIKE '%Boot%' OR C.COURSE_NAME LIKE '%Maker Space%')";
                        sql += " AND sg.storecode IN ('Q1') AND SG.SCHOOLSDCID = 5 AND S.STUDENT_NUMBER =" + stnumb[a] + "";


                        OracleCommand cmdV3 = new OracleCommand(sql, con);
                        cmdV3.ExecuteNonQuery();

                        //EXPLORATORY CONDUCT
                        sql = "CREATE OR REPLACE VIEW MS_PRO_EXP_CON";
                        sql += " AS SELECT S.ID AS STID,SEC.ID AS SECID,S.STUDENT_NUMBER,s.Last_Name,s.First_name,C.COURSE_NAME,SG.STANDARDGRADE AS COND  FROM standardgradesection sg";
                        sql += " LEFT JOIN STANDARD D ON sg.standardid=D.standardid";
                        sql += " LEFT JOIN SECTIONS SEC ON SG.SECTIONSDCID=SEC.dcid";
                        sql += " LEFT JOIN COURSES C ON SEC.COURSE_NUMBER=C.COURSE_NUMBER";
                        sql += " LEFT JOIN TEACHERS T ON SEC.TEACHER=T.id";
                        sql += " LEFT JOIN STUDENTS s ON sg.studentsdcid=s.dcid";
                        sql += " WHERE D.identifier LIKE '%CON%'  AND sg.yearid=28 AND ( C.COURSE_NAME LIKE '%Explora%' OR C.COURSE_NAME LIKE '%Boot%' OR C.COURSE_NAME LIKE '%Maker Space%')";
                        sql += " AND sg.storecode IN ('Q1') AND SG.SCHOOLSDCID=5 AND S.STUDENT_NUMBER=" + stnumb[a] + "";


                        OracleCommand cmdV4 = new OracleCommand(sql, con);
                        cmdV4.ExecuteNonQuery();

                        //EXPLORATORY GRADES
                        sql = "WITH main_query AS(";
                        sql += " SELECT S.STUDENT_NUMBER, T.FIRST_NAME||' '||T.LAST_NAME AS TEACHER, sg.storecode, C.COURSE_NAME, SG.STANDARDGRADE AS EXPLO,";
                        sql += " R.RES, CN.COND, CO.CURRENTABSENCES, CO.CURRENTTARDIES, TO_CHAR(PG.COMMENT_VALUE) AS COMMENTS FROM standardgradesection sg";
                        sql += " LEFT JOIN STANDARD D ON sg.standardid = D.standardid";
                        sql += " LEFT JOIN SECTIONS SEC ON SG.SECTIONSDCID = SEC.dcid";
                        sql += " LEFT JOIN COURSES C ON SEC.COURSE_NUMBER = C.COURSE_NUMBER";
                        sql += " LEFT JOIN TEACHERS T ON SEC.TEACHER = T.id";
                        sql += " LEFT JOIN STUDENTS s ON sg.studentsdcid = s.dcid";
                        sql += " LEFT JOIN MS_PRO_EXP_RES R ON S.ID = R.STID AND SEC.ID = R.SECID";
                        sql += " LEFT JOIN MS_PRO_EXP_CON CN ON S.ID = CN.STID AND SEC.ID = CN.SECID";
                        sql += " LEFT JOIN CC CO ON S.ID = CO.STUDENTID AND SEC.ID = CO.SECTIONID";
                        sql += " LEFT JOIN PGFINALGRADES PG ON S.ID = PG.STUDENTID AND SEC.ID = PG.SECTIONID AND FINALGRADENAME='Q1'";
                        sql += " WHERE D.identifier LIKE '%EXP.%' AND sg.yearid = 28 AND sg.storecode IN ('Q1') AND SG.SCHOOLSDCID = 5 AND S.STUDENT_NUMBER =" + stnumb[a] + "";
                        sql += " )";
                        sql += " SELECT DISTINCT COURSE_NAME,TEACHER";
                        sql += " ,(SELECT  y.EXPLO FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.storecode = 'Q1') Engagement";
                        sql += " ,(SELECT  y.RES FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.storecode = 'Q1') RESP";
                        sql += " ,(SELECT  y.COND FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.storecode = 'Q1') CONDT";
                        sql += " ,(SELECT  y.CURRENTABSENCES FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.storecode = 'Q1') ABS1";
                        sql += " ,(SELECT  y.CURRENTTARDIES FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.storecode = 'Q1') TARDI";
                        sql += " ,(SELECT  y.COMMENTS FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.storecode = 'Q1') Comments";
                        sql += " FROM main_query M";
                        sql += " ORDER BY COURSE_NAME";

                        OracleCommand cmd4 = new OracleCommand(sql, con);
                        OracleDataReader odr4 = cmd4.ExecuteReader();
                        while (odr4.Read())
                        {
                            EXPDATA += odr4["COURSE_NAME"].ToString() + '|';
                            EXPDATA += odr4["TEACHER"].ToString() + '|';
                            EXPDATA += odr4["Engagement"].ToString() + '|';
                            EXPDATA += odr4["RESP"].ToString() + '|';
                            EXPDATA += odr4["CONDT"].ToString() + '|';
                            EXPDATA += odr4["ABS1"].ToString() + '|';
                            EXPDATA += odr4["TARDI"].ToString() + '|';
                            EXPDATA += odr4["Comments"].ToString() + '^';

                        }


                        if (Q1DATA != "")
                        {


                            var stTable = Q1DATA.Split('^');
                            var stAdv = Q1AD.Split('|');
                            var expTable = EXPDATA.Split('^');

                            iTextSharp.text.Image Imagen = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("~/img/WLOGO.jpg"));
                            // Imagen.SetAbsolutePosition(-3, 520);
                            Imagen.ScalePercent(2.5f);


                            PdfPTable HeadT = new PdfPTable(8);
                            HeadT.HorizontalAlignment = Element.ALIGN_CENTER;
                            HeadT.WidthPercentage = 100;

                            PdfPCell logo = new PdfPCell(Imagen);
                            logo.Colspan = 4;
                            logo.Border = 0;
                            logo.HorizontalAlignment = Element.ALIGN_LEFT;
                            logo.Rowspan = 3;
                            logo.Padding = 3;
                            HeadT.AddCell(logo);


                            PdfPCell HS = new PdfPCell(new Phrase("MIDDLE SCHOOL Progress Report", new Font(Font.FontFamily.HELVETICA, 14, Font.BOLD, new BaseColor(135, 0, 27))));
                            HS.HorizontalAlignment = Element.ALIGN_BOTTOM;
                            HS.Colspan = 4;
                            HS.Border = 0;
                            HeadT.AddCell(HS);

                            PdfPCell SQ1 = new PdfPCell(new Phrase("School Year 2018-19 Midsemester 1", new Font(Font.FontFamily.HELVETICA, 14, Font.BOLD, BaseColor.BLACK)));
                            SQ1.HorizontalAlignment = Element.ALIGN_BOTTOM;
                            SQ1.Colspan = 4;
                            SQ1.Border = 0;
                            HeadT.AddCell(SQ1);

                            PdfPCell Pub = new PdfPCell(new Phrase("Published " + DateTime.Now.ToString("MMMM dd, yyyy"), new Font(Font.FontFamily.HELVETICA, 12, Font.ITALIC, BaseColor.BLACK)));
                            Pub.Colspan = 4;
                            Pub.HorizontalAlignment = Element.ALIGN_BOTTOM;
                            Pub.Rowspan = 2;
                            Pub.Border = 0;
                            HeadT.AddCell(Pub);

                            PdfPCell bar1 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                            bar1.HorizontalAlignment = Element.ALIGN_LEFT;
                            bar1.Border = 0;
                            bar1.Colspan = 8;
                            bar1.BackgroundColor = new BaseColor(135, 0, 27);
                            HeadT.AddCell(bar1);

                            PdfPCell stinfo = new PdfPCell(new Phrase("Student Name: " + stAdv[4], new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD, BaseColor.BLACK)));
                            stinfo.HorizontalAlignment = Element.ALIGN_LEFT;
                            stinfo.Border = 0;
                            stinfo.Colspan = 4;
                            stinfo.PaddingTop = 5;
                            HeadT.AddCell(stinfo);

                            PdfPCell messag = new PdfPCell(new Phrase("This report describes progress toward grade level learning expectations, identifies successes and provides guidance for improvement.", new Font(Font.FontFamily.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                            messag.HorizontalAlignment = Element.ALIGN_LEFT;
                            messag.Border = 0;
                            messag.Colspan = 4;
                            messag.Rowspan = 2;
                            HeadT.AddCell(messag);

                            PdfPCell grade = new PdfPCell(new Phrase("Grade: " + stAdv[5], new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD, BaseColor.BLACK)));
                            grade.HorizontalAlignment = Element.ALIGN_LEFT;
                            grade.Border = 0;
                            grade.Colspan = 2;
                            HeadT.AddCell(grade);
                            PdfPCell stid = new PdfPCell(new Phrase("StudentID: " + stAdv[3], new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD, BaseColor.WHITE)));
                            stid.HorizontalAlignment = Element.ALIGN_LEFT;
                            stid.Border = 0;
                            stid.Colspan = 2;
                            stid.PaddingBottom = 5;
                            HeadT.AddCell(stid);

                            PdfPCell boh = new PdfPCell(new Phrase("Advisory: " + stAdv[1] + " (R:" + stAdv[2] + ")", new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD, BaseColor.BLACK)));
                            boh.HorizontalAlignment = Element.ALIGN_LEFT;
                            boh.Border = 0;
                            boh.Colspan = 4;
                            HeadT.AddCell(boh);

                            PdfPCell legdS = new PdfPCell(new Phrase("S = Subject Area Skills       A = Absences" + Environment.NewLine + "R = Responsibility               T = Tardies" + Environment.NewLine + "C = Conduct                        U = Understanding", new Font(Font.FontFamily.HELVETICA, 9, Font.NORMAL, BaseColor.BLACK)));
                            legdS.HorizontalAlignment = Element.ALIGN_LEFT;
                            legdS.Border = 0;
                            legdS.Colspan = 3;
                            legdS.Rowspan = 2;
                            legdS.PaddingBottom = 10;
                            HeadT.AddCell(legdS);

                            PdfPCell LINK = new PdfPCell();
                            var c = new Chunk("Standards Score" + Environment.NewLine + "Assessment Key", new Font(Font.FontFamily.HELVETICA, 9, Font.UNDERLINE, BaseColor.BLUE));
                            c.SetAnchor("http://bit.ly/MSStdAssessKey");
                            LINK.AddElement(c);
                            LINK.Border = 0;
                            LINK.Rowspan = 2;
                            HeadT.AddCell(LINK);



                            PdfPTable GradeTable = new PdfPTable(18);
                            GradeTable.HorizontalAlignment = Element.ALIGN_CENTER;
                            GradeTable.WidthPercentage = 100;

                            PdfPCell Course = new PdfPCell(new Phrase("Course", new Font(Font.FontFamily.HELVETICA, 12.0F, Font.BOLD, BaseColor.WHITE)));
                            Course.HorizontalAlignment = Element.ALIGN_LEFT;

                            Course.BackgroundColor = new BaseColor(135, 0, 27);
                            Course.BorderWidth = 1F;
                            Course.Colspan = 7;
                            Course.PaddingBottom = 5;
                            GradeTable.AddCell(Course);
                            PdfPCell Teacher = new PdfPCell(new Phrase("Teacher", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                            Teacher.HorizontalAlignment = Element.ALIGN_LEFT;

                            Teacher.BackgroundColor = new BaseColor(135, 0, 27);
                            Teacher.BorderWidth = 1F;
                            Teacher.Colspan = 5;
                            Teacher.PaddingBottom = 5;
                            GradeTable.AddCell(Teacher);
                            PdfPCell Q1 = new PdfPCell(new Phrase("S", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                            Q1.HorizontalAlignment = Element.ALIGN_CENTER;
                            Q1.BackgroundColor = new BaseColor(135, 0, 27);
                            Q1.BorderWidth = 1F;
                            Q1.PaddingBottom = 5;
                            Q1.Colspan = 2;
                            GradeTable.AddCell(Q1);
                            PdfPCell Res = new PdfPCell(new Phrase("R", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                            Res.HorizontalAlignment = Element.ALIGN_CENTER;
                            Res.BackgroundColor = new BaseColor(135, 0, 27);
                            Res.BorderWidth = 1F;
                            Res.PaddingBottom = 5;
                            GradeTable.AddCell(Res);
                            PdfPCell Cond = new PdfPCell(new Phrase("C", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                            Cond.HorizontalAlignment = Element.ALIGN_CENTER;
                            Cond.BackgroundColor = new BaseColor(135, 0, 27);
                            Cond.BorderWidth = 1F;
                            Cond.PaddingBottom = 5;
                            GradeTable.AddCell(Cond);
                            PdfPCell ABS = new PdfPCell(new Phrase("A/T", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                            ABS.HorizontalAlignment = Element.ALIGN_CENTER;
                            ABS.BackgroundColor = new BaseColor(135, 0, 27);
                            ABS.BorderWidth = 1F;
                            ABS.Colspan = 2;
                            ABS.PaddingBottom = 5;
                            GradeTable.AddCell(ABS);



                            PdfPCell CO;
                            PdfPCell TE;
                            PdfPCell QU1;
                            PdfPCell COM1;
                            PdfPCell ABS1;
                            PdfPCell R1;
                            PdfPCell Co1;

                            for (int i = 0; i < stTable.Length - 1; i++)
                            {
                                var nfila = stTable[i].Split('|');

                                CO = new PdfPCell(new Phrase(nfila[0], new Font(Font.FontFamily.HELVETICA, 11F, Font.NORMAL, BaseColor.BLACK)));
                                CO.HorizontalAlignment = Element.ALIGN_LEFT;
                                CO.BorderWidth = 0.5F;
                                CO.Colspan = 7;
                                CO.PaddingBottom = 3;
                                CO.BackgroundColor =new BaseColor(235, 235, 235);
                                CO.BorderColor = BaseColor.GRAY;
                                GradeTable.AddCell(CO);

                                TE = new PdfPCell(new Phrase(nfila[1], new Font(Font.FontFamily.HELVETICA, 11F, Font.NORMAL, BaseColor.BLACK)));
                                TE.HorizontalAlignment = Element.ALIGN_LEFT;
                                TE.BorderWidth = 0.5F;
                                TE.Colspan = 5;
                                TE.PaddingBottom = 3;
                                TE.BackgroundColor = new BaseColor(235, 235, 235);
                                TE.BorderColor = BaseColor.GRAY;
                                GradeTable.AddCell(TE);

                                QU1 = new PdfPCell(new Phrase(nfila[2], new Font(Font.FontFamily.HELVETICA, 11F, Font.NORMAL, BaseColor.BLACK)));
                                QU1.HorizontalAlignment = Element.ALIGN_CENTER;
                                QU1.BorderWidth = 0.5F;
                                QU1.Colspan = 2;
                                QU1.PaddingBottom = 3;
                                QU1.BackgroundColor = new BaseColor(235, 235, 235);
                                QU1.BorderColor = BaseColor.GRAY;
                                GradeTable.AddCell(QU1);

                                R1 = new PdfPCell(new Phrase(nfila[3], new Font(Font.FontFamily.HELVETICA, 11F, Font.NORMAL, BaseColor.BLACK)));
                                R1.HorizontalAlignment = Element.ALIGN_CENTER;
                                R1.BorderWidth = 0.5F;
                                R1.PaddingBottom = 3;
                                R1.BackgroundColor = new BaseColor(235, 235, 235);
                                R1.BorderColor = BaseColor.GRAY;
                                GradeTable.AddCell(R1);

                                Co1 = new PdfPCell(new Phrase(nfila[4], new Font(Font.FontFamily.HELVETICA, 11F, Font.NORMAL, BaseColor.BLACK)));
                                Co1.HorizontalAlignment = Element.ALIGN_CENTER;
                                Co1.BorderWidth = 0.5F;
                                Co1.PaddingBottom = 3;
                                Co1.BackgroundColor = new BaseColor(235, 235, 235);
                                Co1.BorderColor = BaseColor.GRAY;
                                GradeTable.AddCell(Co1);

                                ABS1 = new PdfPCell(new Phrase(nfila[5] + '/' + nfila[6], new Font(Font.FontFamily.HELVETICA, 11F, Font.NORMAL, BaseColor.BLACK)));
                                ABS1.HorizontalAlignment = Element.ALIGN_CENTER;
                                ABS1.BorderWidth = 0.5F;
                                ABS1.Colspan = 2;
                                ABS1.PaddingBottom = 3;
                                ABS1.BackgroundColor = new BaseColor(235, 235, 235);
                                ABS1.BorderColor = BaseColor.GRAY;
                                GradeTable.AddCell(ABS1);

                                COM1 = new PdfPCell(new Phrase(nfila[7], new Font(Font.FontFamily.HELVETICA, 10F, Font.NORMAL, BaseColor.BLACK)));
                                COM1.HorizontalAlignment = Element.ALIGN_LEFT;
                                COM1.BorderWidth = 0;
                                COM1.Colspan = 18;
                                COM1.PaddingBottom = 15;
                                //COM1.BorderColor = BaseColor.LIGHT_GRAY;
                                GradeTable.AddCell(COM1);

                            }

                            PdfPTable ExpTable = new PdfPTable(18);
                            ExpTable.HorizontalAlignment = Element.ALIGN_CENTER;
                            ExpTable.WidthPercentage = 100;

                            Paragraph spacio = new Paragraph(" ");

                            PdfPCell expCourse = new PdfPCell(new Phrase("Exploratory", new Font(Font.FontFamily.HELVETICA, 12.0F, Font.BOLD, BaseColor.WHITE)));
                            expCourse.HorizontalAlignment = Element.ALIGN_LEFT;
                            expCourse.BackgroundColor = new BaseColor(135, 0, 27);
                            expCourse.BorderWidth = 1F;
                            expCourse.Colspan = 7;
                            expCourse.PaddingBottom = 5;
                            ExpTable.AddCell(expCourse);
                            PdfPCell expTeacher = new PdfPCell(new Phrase("Teacher", new Font(Font.FontFamily.HELVETICA, 12.0F, Font.BOLD, BaseColor.WHITE)));
                            expTeacher.HorizontalAlignment = Element.ALIGN_LEFT;
                            expTeacher.BackgroundColor = new BaseColor(135, 0, 27);
                            expTeacher.BorderWidth = 1F;
                            expTeacher.Colspan = 5;
                            expTeacher.PaddingBottom = 5;
                            ExpTable.AddCell(expTeacher);
                            PdfPCell expEng = new PdfPCell(new Phrase("U", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                            expEng.HorizontalAlignment = Element.ALIGN_CENTER;
                            expEng.BackgroundColor = new BaseColor(135, 0, 27);
                            expEng.BorderWidth = 1F;
                            expEng.Colspan = 2;
                            expEng.PaddingBottom = 5;
                            ExpTable.AddCell(expEng);
                            PdfPCell expRes = new PdfPCell(new Phrase("R", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                            expRes.HorizontalAlignment = Element.ALIGN_CENTER;
                            expRes.BackgroundColor = new BaseColor(135, 0, 27);
                            expRes.BorderWidth = 1F;
                            expRes.PaddingBottom = 5;
                            ExpTable.AddCell(expRes);
                            PdfPCell expCond = new PdfPCell(new Phrase("C", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                            expCond.HorizontalAlignment = Element.ALIGN_CENTER;
                            expCond.BackgroundColor = new BaseColor(135, 0, 27);
                            expCond.BorderWidth = 1F;
                            expCond.PaddingBottom = 5;
                            ExpTable.AddCell(expCond);
                            PdfPCell expABS = new PdfPCell(new Phrase("A/T", new Font(Font.FontFamily.HELVETICA, 12.0F, Font.BOLD, BaseColor.WHITE)));
                            expABS.HorizontalAlignment = Element.ALIGN_CENTER;
                            expABS.BackgroundColor = new BaseColor(135, 0, 27);
                            expABS.BorderWidth = 1F;
                            expABS.Colspan = 2;
                            expABS.PaddingBottom = 5;
                            ExpTable.AddCell(expABS);

                            PdfPCell expCO;
                            PdfPCell expTE;
                            PdfPCell expQU1;
                            PdfPCell expCOM1;
                            PdfPCell expABS1;
                            PdfPCell expR1;
                            PdfPCell expCo1;

                            for (int i = 0; i < expTable.Length - 1; i++)
                            {
                                var nfila = expTable[i].Split('|');

                                expCO = new PdfPCell(new Phrase(nfila[0], new Font(Font.FontFamily.HELVETICA, 11F, Font.NORMAL, BaseColor.BLACK)));
                                expCO.HorizontalAlignment = Element.ALIGN_LEFT;
                                expCO.BorderWidth = 0.5F;
                                expCO.Colspan = 7;
                                expCO.Padding = 5;
                                expCO.BackgroundColor = new BaseColor(235, 235, 235);
                                expCO.BorderColor = BaseColor.GRAY;
                                ExpTable.AddCell(expCO);

                                expTE = new PdfPCell(new Phrase(nfila[1], new Font(Font.FontFamily.HELVETICA, 11F, Font.NORMAL, BaseColor.BLACK)));
                                expTE.HorizontalAlignment = Element.ALIGN_LEFT;
                                expTE.BorderWidth = 0.5F;
                                expTE.Colspan = 5;
                                expTE.PaddingTop = 5;
                                expTE.BackgroundColor = new BaseColor(235, 235, 235);
                                expTE.BorderColor = BaseColor.GRAY;
                                ExpTable.AddCell(expTE);

                                expQU1 = new PdfPCell(new Phrase(nfila[2], new Font(Font.FontFamily.HELVETICA, 11F, Font.NORMAL, BaseColor.BLACK)));
                                expQU1.HorizontalAlignment = Element.ALIGN_CENTER;
                                expQU1.BorderWidth = 0.5F;
                                expQU1.Colspan = 2;
                                expQU1.PaddingTop = 5;
                                expQU1.BackgroundColor = new BaseColor(235, 235, 235);
                                expQU1.BorderColor = BaseColor.GRAY;
                                ExpTable.AddCell(expQU1);

                                expR1 = new PdfPCell(new Phrase(nfila[3], new Font(Font.FontFamily.HELVETICA, 11F, Font.NORMAL, BaseColor.BLACK)));
                                expR1.HorizontalAlignment = Element.ALIGN_CENTER;
                                expR1.BorderWidth = 0.5F;
                                expR1.PaddingTop = 5;
                                expR1.BackgroundColor = new BaseColor(235, 235, 235);
                                expR1.BorderColor = BaseColor.GRAY;
                                ExpTable.AddCell(expR1);

                                expCo1 = new PdfPCell(new Phrase(nfila[4], new Font(Font.FontFamily.HELVETICA, 11F, Font.NORMAL, BaseColor.BLACK)));
                                expCo1.HorizontalAlignment = Element.ALIGN_CENTER;
                                expCo1.BorderWidth = 0.5F;
                                expCo1.PaddingTop = 5;
                                expCo1.BackgroundColor = new BaseColor(235, 235, 235);
                                expCo1.BorderColor = BaseColor.GRAY;
                                ExpTable.AddCell(expCo1);

                                expABS1 = new PdfPCell(new Phrase(nfila[5] + '/' + nfila[6], new Font(Font.FontFamily.HELVETICA, 11F, Font.NORMAL, BaseColor.BLACK)));
                                expABS1.HorizontalAlignment = Element.ALIGN_CENTER;
                                expABS1.BorderWidth = 0.5F;
                                expABS1.Colspan = 2;
                                expABS1.PaddingTop = 5;
                                expABS1.BackgroundColor = new BaseColor(235, 235, 235);
                                expABS1.BorderColor = BaseColor.GRAY;
                                ExpTable.AddCell(expABS1);

                                expCOM1 = new PdfPCell(new Phrase(nfila[7], new Font(Font.FontFamily.HELVETICA, 11F, Font.NORMAL, BaseColor.BLACK)));
                                expCOM1.HorizontalAlignment = Element.ALIGN_LEFT;
                                expCOM1.BorderWidth = 0;
                                expCOM1.Colspan = 18;
                                expCOM1.PaddingBottom = 15;
                                //expCOM1.BorderColor = BaseColor.LIGHT_GRAY;
                                ExpTable.AddCell(expCOM1);


                            }


                            //documento.Add(stfoto);
                            documento.Add(HeadT);
                            documento.Add(GradeTable);
                            documento.Add(spacio);
                            documento.Add(ExpTable);

                            //Process prc = new System.Diagnostics.Process();
                            //prc.StartInfo.FileName = fileName;
                            //prc.Start();
                        }
                        else
                        {
                            con.Close();

                        }
                        documento.NewPage();
                    }

                    con.Close();

                }
                else
                {
                    if (con.State == ConnectionState.Closed)
                    {
                        con.Open();
                    }
                    fname = "MS_ProgressReport_" + DateTime.Now.DayOfYear + DateTime.Now.Month + DateTime.Now.Day + DateTime.Now.Millisecond + ".pdf";
                    fileName = HttpContext.Current.Server.MapPath("~/RepoFiles/" + fname);
                    PdfWriter.GetInstance(documento, new FileStream(fileName, FileMode.Create));
                    documento.Open();

                    string Q1DATA = string.Empty;
                    string Q1AD = string.Empty;
                    string EXPDATA = string.Empty;
                    if (con.State == ConnectionState.Closed)
                    {
                        con.Open();
                    }

                    //CREATE OR MODIFY VIEWS

                    sql = "CREATE OR REPLACE VIEW MS_PRO_REP_RES";
                    sql += " AS SELECT S.ID,SEC.ID AS SECTID,S.STUDENT_NUMBER,s.Last_Name,s.First_name,C.COURSE_NAME,SG.STANDARDGRADE AS RES FROM standardgradesection sg";
                    sql += " LEFT JOIN STANDARD D ON sg.standardid = D.standardid";
                    sql += " LEFT JOIN STANDARD Di ON sg.standardid = Di.standardid";
                    sql += " LEFT JOIN SECTIONS SEC ON SG.SECTIONSDCID = SEC.dcid";
                    sql += " LEFT JOIN COURSES C ON SEC.COURSE_NUMBER = C.COURSE_NUMBER";
                    sql += " LEFT JOIN TEACHERS T ON SEC.TEACHER = T.id";
                    sql += " LEFT JOIN STUDENTS s ON sg.studentsdcid = s.dcid";
                    sql += " WHERE D.identifier LIKE '%RES%' AND C.COURSE_NAME NOT LIKE '%Explora%' AND C.COURSE_NAME NOT LIKE '%Advisory%' AND C.COURSE_NAME NOT LIKE '%Space%' AND C.COURSE_NAME NOT LIKE '%Boot%' AND sg.yearid = 28";
                    sql += " AND sg.storecode IN ('Q1') AND SG.SCHOOLSDCID = 5 AND S.STUDENT_NUMBER =" + stnum + "";


                    OracleCommand cmdV1 = new OracleCommand(sql, con);
                    cmdV1.ExecuteNonQuery();

                    //CONDUCT.
                    sql = "CREATE OR REPLACE VIEW MS_PRO_REP_CON";
                    sql += " AS SELECT S.ID,SEC.ID AS SECID,S.STUDENT_NUMBER,s.Last_Name,s.First_name,C.COURSE_NAME,SG.STANDARDGRADE AS COND FROM standardgradesection sg";
                    sql += " LEFT JOIN STANDARD D ON sg.standardid = D.standardid";
                    sql += " LEFT JOIN SECTIONS SEC ON SG.SECTIONSDCID = SEC.dcid";
                    sql += " LEFT JOIN COURSES C ON SEC.COURSE_NUMBER = C.COURSE_NUMBER";
                    sql += " LEFT JOIN TEACHERS T ON SEC.TEACHER = T.id";
                    sql += " LEFT JOIN STUDENTS s ON sg.studentsdcid = s.dcid";
                    sql += " WHERE D.identifier LIKE '%CON%' AND C.COURSE_NAME NOT LIKE '%Explora%' AND C.COURSE_NAME NOT LIKE '%Advisory%' AND C.COURSE_NAME NOT LIKE '%Space%' AND C.COURSE_NAME NOT LIKE '%Boot%' AND sg.yearid = 28";
                    sql += " AND sg.storecode IN ('Q1') AND SG.SCHOOLSDCID = 5 AND S.STUDENT_NUMBER = " + stnum + "";


                    OracleCommand cmdV2 = new OracleCommand(sql, con);
                    cmdV2.ExecuteNonQuery();


                    //Advisory Teacher
                    sql = "SELECT DISTINCT C.COURSE_NAME,T.FIRST_NAME||' '||T.LAST_NAME AS TEACHER,R.RES,S.STUDENT_NUMBER,S.FIRST_NAME||' '||S.LAST_NAME AS STUDENT,S.GRADE_LEVEL FROM CC CO";
                    sql += " LEFT JOIN STUDENTS S ON CO.STUDENTID = S.ID";
                    sql += " LEFT JOIN COURSES C ON CO.COURSE_NUMBER = C.COURSE_NUMBER";
                    sql += " LEFT JOIN TEACHERS T ON CO.TEACHERID = T.ID";
                    sql += " LEFT JOIN MS_PRO_REP_RES R ON CO.STUDENTID = R.ID AND CO.SECTIONID = R.SECTID";
                    sql += " WHERE CO.TERMID IN(2800, 2801, 2802)  AND C.COURSE_NAME LIKE '%Advisory%' AND S.STUDENT_NUMBER =" + stnum + "";

                    OracleCommand cmd = new OracleCommand(sql, con);
                    OracleDataReader odr = cmd.ExecuteReader();
                    while (odr.Read())
                    {
                        Q1AD += odr["COURSE_NAME"].ToString() + '|';
                        Q1AD += odr["TEACHER"].ToString() + '|';
                        Q1AD += odr["RES"].ToString() + '|';
                        Q1AD += odr["STUDENT_NUMBER"].ToString() + '|';
                        Q1AD += odr["STUDENT"].ToString() + '|';
                        Q1AD += odr["GRADE_LEVEL"].ToString() + '|';

                    }


                    //GRADES VALUES
                    sql = "WITH main_query AS ( SELECT C.COURSE_NAME,T.FIRST_NAME||' '||T.LAST_NAME AS TEACHER, PG.FINALGRADENAME,PG.GRADE,R.RES,CN.COND,";
                    sql += " TO_CHAR(PG.COMMENT_VALUE) AS COMMENTS,CO.CURRENTABSENCES,CO.CURRENTTARDIES FROM CC CO";
                    sql += " LEFT JOIN STUDENTS S ON CO.STUDENTID=S.ID";
                    sql += " LEFT JOIN COURSES C ON CO.COURSE_NUMBER=C.COURSE_NUMBER";
                    sql += " LEFT JOIN TEACHERS T ON CO.TEACHERID=T.ID";
                    sql += " LEFT JOIN PGFINALGRADES PG ON CO.STUDENTID=PG.STUDENTID AND CO.SECTIONID=PG.SECTIONID";
                    sql += " LEFT JOIN MS_PRO_REP_RES R ON CO.STUDENTID = R.ID AND CO.SECTIONID = R.SECTID";
                    sql += " LEFT JOIN MS_PRO_REP_CON CN ON CO.STUDENTID = CN.ID AND CO.SECTIONID = CN.SECID";
                    sql += " WHERE CO.TERMID IN(2800,2801,2802) AND PG.FINALGRADENAME IN ('Q1') AND C.COURSE_NAME NOT LIKE '%Explora%' AND C.COURSE_NAME NOT LIKE '%Space%' AND C.COURSE_NAME NOT LIKE '%Advisory%' AND C.COURSE_NAME NOT LIKE '%Boot%' AND S.STUDENT_NUMBER=" + stnum + "";
                    sql += " )";
                    sql += " SELECT DISTINCT COURSE_NAME,TEACHER";
                    sql += " ,(SELECT  y.GRADE FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.FINALGRADENAME='Q1') SKILL";
                    sql += " ,(SELECT  y.RES FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.FINALGRADENAME='Q1') RESP";
                    sql += " ,(SELECT  y.COND FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.FINALGRADENAME='Q1') CONDT";
                    sql += " ,(SELECT  y.CURRENTABSENCES FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.FINALGRADENAME='Q1') ABS1";
                    sql += " ,(SELECT  y.CURRENTTARDIES FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.FINALGRADENAME = 'Q1') TARDI";
                    sql += " ,(SELECT  y.COMMENTS FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.FINALGRADENAME = 'Q1') Comments";
                    sql += "  FROM	main_query M";
                    sql += " ORDER BY COURSE_NAME";

                    OracleCommand cmd1 = new OracleCommand(sql, con);
                    OracleDataReader odr1 = cmd1.ExecuteReader();
                    while (odr1.Read())
                    {
                        Q1DATA += odr1["COURSE_NAME"].ToString() + '|';
                        Q1DATA += odr1["TEACHER"].ToString() + '|';
                        Q1DATA += odr1["SKILL"].ToString() + '|';
                        Q1DATA += odr1["RESP"].ToString() + '|';
                        Q1DATA += odr1["CONDT"].ToString() + '|';
                        Q1DATA += odr1["ABS1"].ToString() + '|';
                        Q1DATA += odr1["TARDI"].ToString() + '|';
                        Q1DATA += odr1["Comments"].ToString() + '^';

                    }

                    //EXPLORATORY RES.
                    sql = "CREATE OR REPLACE VIEW MS_PRO_EXP_RES";
                    sql += " AS SELECT S.ID AS STID,SEC.ID AS SECID,S.STUDENT_NUMBER,s.Last_Name,s.First_name,C.COURSE_NAME,SG.STANDARDGRADE AS RES FROM standardgradesection sg";
                    sql += " LEFT JOIN STANDARD D ON sg.standardid = D.standardid";
                    sql += " LEFT JOIN SECTIONS SEC ON SG.SECTIONSDCID = SEC.dcid";
                    sql += " LEFT JOIN COURSES C ON SEC.COURSE_NUMBER = C.COURSE_NUMBER";
                    sql += " LEFT JOIN TEACHERS T ON SEC.TEACHER = T.id";
                    sql += " LEFT JOIN STUDENTS s ON sg.studentsdcid = s.dcid";
                    sql += " WHERE D.identifier LIKE '%RES%' AND sg.yearid = 28 AND";
                    sql += " (C.COURSE_NAME LIKE '%Explora%' OR C.COURSE_NAME LIKE '%Boot%' OR C.COURSE_NAME LIKE '%Maker Space%')";
                    sql += " AND sg.storecode IN ('Q1') AND SG.SCHOOLSDCID = 5 AND S.STUDENT_NUMBER =" + stnum + "";


                    OracleCommand cmdV3 = new OracleCommand(sql, con);
                    cmdV3.ExecuteNonQuery();

                    //EXPLORATORY CONDUCT
                    sql = "CREATE OR REPLACE VIEW MS_PRO_EXP_CON";
                    sql += " AS SELECT S.ID AS STID,SEC.ID AS SECID,S.STUDENT_NUMBER,s.Last_Name,s.First_name,C.COURSE_NAME,SG.STANDARDGRADE AS COND  FROM standardgradesection sg";
                    sql += " LEFT JOIN STANDARD D ON sg.standardid=D.standardid";
                    sql += " LEFT JOIN SECTIONS SEC ON SG.SECTIONSDCID=SEC.dcid";
                    sql += " LEFT JOIN COURSES C ON SEC.COURSE_NUMBER=C.COURSE_NUMBER";
                    sql += " LEFT JOIN TEACHERS T ON SEC.TEACHER=T.id";
                    sql += " LEFT JOIN STUDENTS s ON sg.studentsdcid=s.dcid";
                    sql += " WHERE D.identifier LIKE '%CON%'  AND sg.yearid=28 AND ( C.COURSE_NAME LIKE '%Explora%' OR C.COURSE_NAME LIKE '%Boot%' OR C.COURSE_NAME LIKE '%Maker Space%')";
                    sql += " AND sg.storecode IN ('Q1') AND SG.SCHOOLSDCID=5 AND S.STUDENT_NUMBER=" + stnum + "";


                    OracleCommand cmdV4 = new OracleCommand(sql, con);
                    cmdV4.ExecuteNonQuery();

                    //EXPLORATORY GRADES
                    sql = "WITH main_query AS(";
                    sql += " SELECT S.STUDENT_NUMBER, T.FIRST_NAME||' '||T.LAST_NAME AS TEACHER, sg.storecode, C.COURSE_NAME, SG.STANDARDGRADE AS EXPLO,";
                    sql += " R.RES, CN.COND, CO.CURRENTABSENCES, CO.CURRENTTARDIES, TO_CHAR(PG.COMMENT_VALUE) AS COMMENTS FROM standardgradesection sg";
                    sql += " LEFT JOIN STANDARD D ON sg.standardid = D.standardid";
                    sql += " LEFT JOIN SECTIONS SEC ON SG.SECTIONSDCID = SEC.dcid";
                    sql += " LEFT JOIN COURSES C ON SEC.COURSE_NUMBER = C.COURSE_NUMBER";
                    sql += " LEFT JOIN TEACHERS T ON SEC.TEACHER = T.id";
                    sql += " LEFT JOIN STUDENTS s ON sg.studentsdcid = s.dcid";
                    sql += " LEFT JOIN MS_PRO_EXP_RES R ON S.ID = R.STID AND SEC.ID = R.SECID";
                    sql += " LEFT JOIN MS_PRO_EXP_CON CN ON S.ID = CN.STID AND SEC.ID = CN.SECID";
                    sql += " LEFT JOIN CC CO ON S.ID = CO.STUDENTID AND SEC.ID = CO.SECTIONID";
                    sql += " LEFT JOIN PGFINALGRADES PG ON S.ID = PG.STUDENTID AND SEC.ID = PG.SECTIONID AND FINALGRADENAME='Q1'";
                    sql += " WHERE D.identifier LIKE '%EXP.%' AND sg.yearid = 28 AND sg.storecode IN ('Q1') AND SG.SCHOOLSDCID = 5 AND S.STUDENT_NUMBER =" + stnum + "";
                    sql += " )";
                    sql += " SELECT DISTINCT COURSE_NAME,TEACHER";
                    sql += " ,(SELECT  y.EXPLO FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.storecode = 'Q1') Engagement";
                    sql += " ,(SELECT  y.RES FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.storecode = 'Q1') RESP";
                    sql += " ,(SELECT  y.COND FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.storecode = 'Q1') CONDT";
                    sql += " ,(SELECT  y.CURRENTABSENCES FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.storecode = 'Q1') ABS1";
                    sql += " ,(SELECT  y.CURRENTTARDIES FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.storecode = 'Q1') TARDI";
                    sql += " ,(SELECT  y.COMMENTS FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.storecode = 'Q1') Comments";
                    sql += " FROM main_query M";
                    sql += " ORDER BY COURSE_NAME";

                    OracleCommand cmd4 = new OracleCommand(sql, con);
                    OracleDataReader odr4 = cmd4.ExecuteReader();
                    while (odr4.Read())
                    {
                        EXPDATA += odr4["COURSE_NAME"].ToString() + '|';
                        EXPDATA += odr4["TEACHER"].ToString() + '|';
                        EXPDATA += odr4["Engagement"].ToString() + '|';
                        EXPDATA += odr4["RESP"].ToString() + '|';
                        EXPDATA += odr4["CONDT"].ToString() + '|';
                        EXPDATA += odr4["ABS1"].ToString() + '|';
                        EXPDATA += odr4["TARDI"].ToString() + '|';
                        EXPDATA += odr4["Comments"].ToString() + '^';

                    }


                    if (Q1DATA != "")
                    {


                        var stTable = Q1DATA.Split('^');
                        var stAdv = Q1AD.Split('|');
                        var expTable = EXPDATA.Split('^');
                        iTextSharp.text.Image Imagen = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("~/img/WLOGO.jpg"));
                        // Imagen.SetAbsolutePosition(-3, 520);
                        Imagen.ScalePercent(2.5f);


                        PdfPTable HeadT = new PdfPTable(8);
                        HeadT.HorizontalAlignment = Element.ALIGN_CENTER;
                        HeadT.WidthPercentage = 100;

                        PdfPCell logo = new PdfPCell(Imagen);
                        logo.Colspan = 4;
                        logo.Border = 0;
                        logo.HorizontalAlignment = Element.ALIGN_LEFT;
                        logo.Rowspan = 3;
                        logo.Padding = 3;
                        HeadT.AddCell(logo);


                        PdfPCell HS = new PdfPCell(new Phrase("MIDDLE SCHOOL Progress Report", new Font(Font.FontFamily.HELVETICA, 14, Font.BOLD, new BaseColor(135, 0, 27))));
                        HS.HorizontalAlignment = Element.ALIGN_BOTTOM;
                        HS.Colspan = 4;
                        HS.Border = 0;
                        HeadT.AddCell(HS);

                        PdfPCell SQ1 = new PdfPCell(new Phrase("School Year 2018-19 Midsemester 1", new Font(Font.FontFamily.HELVETICA, 14, Font.BOLD, BaseColor.BLACK)));
                        SQ1.HorizontalAlignment = Element.ALIGN_BOTTOM;
                        SQ1.Colspan = 4;
                        SQ1.Border = 0;
                        HeadT.AddCell(SQ1);

                        PdfPCell Pub = new PdfPCell(new Phrase("Published " + DateTime.Now.ToString("MMMM dd, yyyy"), new Font(Font.FontFamily.HELVETICA, 12, Font.ITALIC, BaseColor.BLACK)));
                        Pub.Colspan = 4;
                        Pub.HorizontalAlignment = Element.ALIGN_BOTTOM;
                        Pub.Rowspan = 2;
                        Pub.Border = 0;
                        HeadT.AddCell(Pub);

                        PdfPCell bar1 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                        bar1.HorizontalAlignment = Element.ALIGN_LEFT;
                        bar1.Border = 0;
                        bar1.Colspan = 8;
                        bar1.BackgroundColor = new BaseColor(135, 0, 27);
                        HeadT.AddCell(bar1);

                        PdfPCell stinfo = new PdfPCell(new Phrase("Student Name: " + stAdv[4], new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD, BaseColor.BLACK)));
                        stinfo.HorizontalAlignment = Element.ALIGN_LEFT;
                        stinfo.Border = 0;
                        stinfo.Colspan = 4;
                        stinfo.PaddingTop = 5;
                        HeadT.AddCell(stinfo);

                        PdfPCell messag = new PdfPCell(new Phrase("This report describes progress toward grade level learning expectations, identifies successes and provides guidance for improvement.", new Font(Font.FontFamily.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                        messag.HorizontalAlignment = Element.ALIGN_LEFT;
                        messag.Border = 0;
                        messag.Colspan = 4;
                        messag.Rowspan = 2;
                        HeadT.AddCell(messag);

                        PdfPCell grade = new PdfPCell(new Phrase("Grade: " + stAdv[5], new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD, BaseColor.BLACK)));
                        grade.HorizontalAlignment = Element.ALIGN_LEFT;
                        grade.Border = 0;
                        grade.Colspan = 2;
                        HeadT.AddCell(grade);
                        PdfPCell stid = new PdfPCell(new Phrase("StudentID: " + stAdv[3], new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD, BaseColor.WHITE)));
                        stid.HorizontalAlignment = Element.ALIGN_LEFT;
                        stid.Border = 0;
                        stid.Colspan = 2;
                        stid.PaddingBottom = 5;
                        HeadT.AddCell(stid);

                        PdfPCell boh = new PdfPCell(new Phrase("Advisory: " + stAdv[1] + " (R:" + stAdv[2] + ")", new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD, BaseColor.BLACK)));
                        boh.HorizontalAlignment = Element.ALIGN_LEFT;
                        boh.Border = 0;
                        boh.Colspan = 4;
                        HeadT.AddCell(boh);

                        PdfPCell legdS = new PdfPCell(new Phrase("S = Subject Area Skills       A = Absences"+Environment.NewLine+ "R = Responsibility               T = Tardies"+ Environment.NewLine+ "C = Conduct                        U = Understanding", new Font(Font.FontFamily.HELVETICA, 9, Font.NORMAL, BaseColor.BLACK)));
                        legdS.HorizontalAlignment = Element.ALIGN_LEFT;
                        legdS.Border = 0;
                        legdS.Colspan = 3;
                        legdS.Rowspan = 2;
                        legdS.PaddingBottom = 10;
                        HeadT.AddCell(legdS);

                        PdfPCell LINK = new PdfPCell();
                        var c = new Chunk("Standards Score"+Environment.NewLine+ "Assessment Key", new Font(Font.FontFamily.HELVETICA, 9, Font.UNDERLINE, BaseColor.BLUE));
                        c.SetAnchor("http://bit.ly/MSStdAssessKey");
                        LINK.AddElement(c);
                        LINK.Border = 0;
                        LINK.Rowspan = 2;
                        HeadT.AddCell(LINK);

                        PdfPTable GradeTable = new PdfPTable(18);
                        GradeTable.HorizontalAlignment = Element.ALIGN_CENTER;
                        GradeTable.WidthPercentage = 100;

                        PdfPCell Course = new PdfPCell(new Phrase("Course", new Font(Font.FontFamily.HELVETICA, 12.0F, Font.BOLD, BaseColor.WHITE)));
                        Course.HorizontalAlignment = Element.ALIGN_LEFT;
                        
                        Course.BackgroundColor = new BaseColor(135, 0, 27);
                        Course.BorderWidth = 1F;
                        Course.Colspan = 7;
                        Course.PaddingBottom = 5;
                        GradeTable.AddCell(Course);
                        PdfPCell Teacher = new PdfPCell(new Phrase("Teacher", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                        Teacher.HorizontalAlignment = Element.ALIGN_LEFT;
                        
                        Teacher.BackgroundColor = new BaseColor(135, 0, 27);
                        Teacher.BorderWidth = 1F;
                        Teacher.Colspan = 5;
                        Teacher.PaddingBottom = 5;
                        GradeTable.AddCell(Teacher);
                        PdfPCell Q1 = new PdfPCell(new Phrase("S", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                        Q1.HorizontalAlignment = Element.ALIGN_CENTER;
                        Q1.BackgroundColor = new BaseColor(135, 0, 27);
                        Q1.BorderWidth = 1F;
                        Q1.PaddingBottom = 5;
                        Q1.Colspan = 2;
                        GradeTable.AddCell(Q1);
                        PdfPCell Res = new PdfPCell(new Phrase("R", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                        Res.HorizontalAlignment = Element.ALIGN_CENTER;
                        Res.BackgroundColor = new BaseColor(135, 0, 27);
                        Res.BorderWidth = 1F;
                        Res.PaddingBottom = 5;
                        GradeTable.AddCell(Res);
                        PdfPCell Cond = new PdfPCell(new Phrase("C", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                        Cond.HorizontalAlignment = Element.ALIGN_CENTER;
                        Cond.BackgroundColor = new BaseColor(135, 0, 27);
                        Cond.BorderWidth = 1F;
                        Cond.PaddingBottom = 5;
                        GradeTable.AddCell(Cond);
                        PdfPCell ABS = new PdfPCell(new Phrase("A/T", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                        ABS.HorizontalAlignment = Element.ALIGN_CENTER;
                        ABS.BackgroundColor = new BaseColor(135, 0, 27);
                        ABS.BorderWidth = 1F;
                        ABS.Colspan = 2;
                        ABS.PaddingBottom = 8;
                        GradeTable.AddCell(ABS);



                        PdfPCell CO;
                        PdfPCell TE;
                        PdfPCell QU1;
                        PdfPCell COM1;
                        PdfPCell ABS1;
                        PdfPCell R1;
                        PdfPCell Co1;

                        for (int i = 0; i < stTable.Length - 1; i++)
                        {
                            var nfila = stTable[i].Split('|');

                            CO = new PdfPCell(new Phrase(nfila[0], new Font(Font.FontFamily.HELVETICA, 11F, Font.NORMAL, BaseColor.BLACK)));
                            CO.HorizontalAlignment = Element.ALIGN_LEFT;
                            CO.BorderWidth = 0.5F;
                            CO.Colspan = 7;
                            CO.PaddingBottom = 3;
                            CO.BackgroundColor = new BaseColor(235, 235, 235);
                            CO.BorderColor = BaseColor.GRAY;
                            GradeTable.AddCell(CO);

                            TE = new PdfPCell(new Phrase(nfila[1], new Font(Font.FontFamily.HELVETICA, 11F, Font.NORMAL, BaseColor.BLACK)));
                            TE.HorizontalAlignment = Element.ALIGN_LEFT;
                            TE.BorderWidth = 0.5F;
                            TE.Colspan = 5;
                            TE.PaddingBottom = 3;
                            TE.BackgroundColor = new BaseColor(235, 235, 235);
                            TE.BorderColor = BaseColor.GRAY;
                            GradeTable.AddCell(TE);

                            QU1 = new PdfPCell(new Phrase(nfila[2], new Font(Font.FontFamily.HELVETICA, 11F, Font.NORMAL, BaseColor.BLACK)));
                            QU1.HorizontalAlignment = Element.ALIGN_CENTER;
                            QU1.BorderWidth = 0.5F;
                            QU1.Colspan = 2;
                            QU1.PaddingBottom = 3;
                            QU1.BackgroundColor = new BaseColor(235, 235, 235);
                            QU1.BorderColor = BaseColor.GRAY;
                            GradeTable.AddCell(QU1);

                            R1 = new PdfPCell(new Phrase(nfila[3], new Font(Font.FontFamily.HELVETICA, 11F, Font.NORMAL, BaseColor.BLACK)));
                            R1.HorizontalAlignment = Element.ALIGN_CENTER;
                            R1.BorderWidth = 0.5F;
                            R1.PaddingBottom = 3;
                            R1.BackgroundColor = new BaseColor(235, 235, 235);
                            R1.BorderColor = BaseColor.GRAY;
                            GradeTable.AddCell(R1);

                            Co1 = new PdfPCell(new Phrase(nfila[4], new Font(Font.FontFamily.HELVETICA, 11F, Font.NORMAL, BaseColor.BLACK)));
                            Co1.HorizontalAlignment = Element.ALIGN_CENTER;
                            Co1.BorderWidth = 0.5F;
                            Co1.PaddingBottom = 3;
                            Co1.BackgroundColor = new BaseColor(235, 235, 235);
                            Co1.BorderColor = BaseColor.GRAY;
                            GradeTable.AddCell(Co1);

                            ABS1 = new PdfPCell(new Phrase(nfila[5] + '/' + nfila[6], new Font(Font.FontFamily.HELVETICA, 11F, Font.NORMAL, BaseColor.BLACK)));
                            ABS1.HorizontalAlignment = Element.ALIGN_CENTER;
                            ABS1.BorderWidth = 0.5F;
                            ABS1.Colspan = 2;
                            ABS1.PaddingBottom = 3;
                            ABS1.BackgroundColor = new BaseColor(235, 235, 235);
                            ABS1.BorderColor = BaseColor.GRAY;
                            GradeTable.AddCell(ABS1);

                            COM1 = new PdfPCell(new Phrase(nfila[7], new Font(Font.FontFamily.HELVETICA, 10F, Font.NORMAL, BaseColor.BLACK)));
                            COM1.HorizontalAlignment = Element.ALIGN_LEFT;
                            COM1.BorderWidth = 0;
                            COM1.Colspan = 18;
                            COM1.PaddingBottom = 15;
                            //COM1.BorderColor = BaseColor.LIGHT_GRAY;
                            GradeTable.AddCell(COM1);

                        }


                        PdfPTable ExpTable = new PdfPTable(18);
                        ExpTable.HorizontalAlignment = Element.ALIGN_CENTER;
                        ExpTable.WidthPercentage = 100;

                        Paragraph spacio = new Paragraph(" ");

                        PdfPCell expCourse = new PdfPCell(new Phrase("Exploratory", new Font(Font.FontFamily.HELVETICA, 12.0F, Font.BOLD, BaseColor.WHITE)));
                        expCourse.HorizontalAlignment = Element.ALIGN_LEFT;
                        expCourse.BackgroundColor = new BaseColor(135, 0, 27);
                        expCourse.BorderWidth = 1F;
                        expCourse.Colspan = 7;
                        expCourse.PaddingBottom = 5;
                        ExpTable.AddCell(expCourse);
                        PdfPCell expTeacher = new PdfPCell(new Phrase("Teacher", new Font(Font.FontFamily.HELVETICA, 12.0F, Font.BOLD, BaseColor.WHITE)));
                        expTeacher.HorizontalAlignment = Element.ALIGN_LEFT;
                        expTeacher.BackgroundColor = new BaseColor(135, 0, 27);
                        expTeacher.BorderWidth = 1F;
                        expTeacher.Colspan = 5;
                        expTeacher.PaddingBottom = 5;
                        ExpTable.AddCell(expTeacher);
                        PdfPCell expEng = new PdfPCell(new Phrase("U", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                        expEng.HorizontalAlignment = Element.ALIGN_CENTER;
                        expEng.BackgroundColor = new BaseColor(135, 0, 27);
                        expEng.BorderWidth = 1F;
                        expEng.Colspan = 2;
                        expEng.PaddingBottom = 5;
                        ExpTable.AddCell(expEng);
                        PdfPCell expRes = new PdfPCell(new Phrase("R", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                        expRes.HorizontalAlignment = Element.ALIGN_CENTER;
                        expRes.BackgroundColor = new BaseColor(135, 0, 27);
                        expRes.BorderWidth = 1F;
                        expRes.PaddingBottom = 5;
                        ExpTable.AddCell(expRes);
                        PdfPCell expCond = new PdfPCell(new Phrase("C", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                        expCond.HorizontalAlignment = Element.ALIGN_CENTER;
                        expCond.BackgroundColor = new BaseColor(135, 0, 27);
                        expCond.BorderWidth = 1F;
                        expCond.PaddingBottom = 5;
                        ExpTable.AddCell(expCond);
                        PdfPCell expABS = new PdfPCell(new Phrase("A/T", new Font(Font.FontFamily.HELVETICA, 12.0F, Font.BOLD, BaseColor.WHITE)));
                        expABS.HorizontalAlignment = Element.ALIGN_CENTER;
                        expABS.BackgroundColor = new BaseColor(135, 0, 27);
                        expABS.BorderWidth = 1F;
                        expABS.Colspan = 2;
                        expABS.PaddingBottom = 5;
                        ExpTable.AddCell(expABS);

                        PdfPCell expCO;
                        PdfPCell expTE;
                        PdfPCell expQU1;
                        PdfPCell expCOM1;
                        PdfPCell expABS1;
                        PdfPCell expR1;
                        PdfPCell expCo1;

                        for (int i = 0; i < expTable.Length - 1; i++)
                        {
                            var nfila = expTable[i].Split('|');

                            expCO = new PdfPCell(new Phrase(nfila[0], new Font(Font.FontFamily.HELVETICA, 11F, Font.NORMAL, BaseColor.BLACK)));
                            expCO.HorizontalAlignment = Element.ALIGN_LEFT;
                            expCO.BorderWidth = 0.5F;
                            expCO.Colspan = 7;
                            expCO.Padding = 5;
                            expCO.BackgroundColor = new BaseColor(235, 235, 235);
                            expCO.BorderColor = BaseColor.GRAY;
                            ExpTable.AddCell(expCO);

                            expTE = new PdfPCell(new Phrase(nfila[1], new Font(Font.FontFamily.HELVETICA, 11F, Font.NORMAL, BaseColor.BLACK)));
                            expTE.HorizontalAlignment = Element.ALIGN_LEFT;
                            expTE.BorderWidth = 0.5F;
                            expTE.Colspan = 5;
                            expTE.PaddingTop = 5;
                            expTE.BackgroundColor = new BaseColor(235, 235, 235);
                            expTE.BorderColor = BaseColor.GRAY;
                            ExpTable.AddCell(expTE);

                            expQU1 = new PdfPCell(new Phrase(nfila[2], new Font(Font.FontFamily.HELVETICA, 11F, Font.NORMAL, BaseColor.BLACK)));
                            expQU1.HorizontalAlignment = Element.ALIGN_CENTER;
                            expQU1.BorderWidth = 0.5F;
                            expQU1.Colspan = 2;
                            expQU1.PaddingTop = 5; 
                            expQU1.BackgroundColor = new BaseColor(235, 235, 235);
                            expQU1.BorderColor = BaseColor.GRAY;
                            ExpTable.AddCell(expQU1);

                            expR1 = new PdfPCell(new Phrase(nfila[3], new Font(Font.FontFamily.HELVETICA, 11F, Font.NORMAL, BaseColor.BLACK)));
                            expR1.HorizontalAlignment = Element.ALIGN_CENTER;
                            expR1.BorderWidth = 0.5F;
                            expR1.PaddingTop = 5;
                            expR1.BackgroundColor = new BaseColor(235, 235, 235);
                            expR1.BorderColor = BaseColor.GRAY;
                            ExpTable.AddCell(expR1);

                            expCo1 = new PdfPCell(new Phrase(nfila[4], new Font(Font.FontFamily.HELVETICA, 11F, Font.NORMAL, BaseColor.BLACK)));
                            expCo1.HorizontalAlignment = Element.ALIGN_CENTER;
                            expCo1.BorderWidth = 0.5F;
                            expCo1.PaddingTop = 5;
                            expCo1.BackgroundColor = new BaseColor(235, 235, 235);
                            expCo1.BorderColor = BaseColor.GRAY;
                            ExpTable.AddCell(expCo1);

                            expABS1 = new PdfPCell(new Phrase(nfila[5] + '/' + nfila[6], new Font(Font.FontFamily.HELVETICA, 11F, Font.NORMAL, BaseColor.BLACK)));
                            expABS1.HorizontalAlignment = Element.ALIGN_CENTER;
                            expABS1.BorderWidth = 0.5F;
                            expABS1.Colspan = 2;
                            expABS1.PaddingTop = 5;
                            expABS1.BackgroundColor = new BaseColor(235, 235, 235);
                            expABS1.BorderColor = BaseColor.GRAY;
                            ExpTable.AddCell(expABS1);

                            expCOM1 = new PdfPCell(new Phrase(nfila[7], new Font(Font.FontFamily.HELVETICA, 10F, Font.NORMAL, BaseColor.BLACK)));
                            expCOM1.HorizontalAlignment = Element.ALIGN_LEFT;
                            expCOM1.BorderWidth = 0;
                            expCOM1.Colspan = 18;
                            expCOM1.PaddingBottom = 15;
                            //expCOM1.BorderColor = BaseColor.LIGHT_GRAY;
                            ExpTable.AddCell(expCOM1);


                        }
                        //documento.Add(stfoto);
                        documento.Add(HeadT);
                        documento.Add(GradeTable);
                        documento.Add(spacio);
                        documento.Add(ExpTable);
                        



                        //Process prc = new System.Diagnostics.Process();
                        //prc.StartInfo.FileName = fileName;
                        //prc.Start();
                    }
                    else
                    {
                        con.Close();
                        fname = "";
                    }
                    con.Close();

                }

                documento.Close();
                con.Dispose();
            }
            catch (Exception ex)
            {
                throw;
            }
            return fname;
        }

        [WebMethod]
        public static string PROGREPORTQ1(string stnum)
        {
            string sql = string.Empty;
            


            string fname = string.Empty;
            string fileName = string.Empty;
            OracleConnection con = new OracleConnection();
            con.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings["conn"].ConnectionString;
            Document documento = new Document(PageSize.LETTER, 10, 10, 5, 5);
            try
            {

                if (stnum.IndexOf(';') > -1)
                {
                    var stnumb = stnum.Split(';');
                    fname = "MS_ProgressReport_" + DateTime.Now.DayOfYear + DateTime.Now.Month + DateTime.Now.Day + DateTime.Now.Millisecond + ".pdf";
                    fileName = HttpContext.Current.Server.MapPath("~/RepoFiles/" + fname);
                    PdfWriter.GetInstance(documento, new FileStream(fileName, FileMode.Create));
                    documento.Open();

                    for (int a = 0; a < stnumb.Length; a++)
                    {
                        string Q1DATA = string.Empty;
                        string Q1AD = string.Empty;

                        if (con.State == ConnectionState.Closed)
                        {
                            con.Open();
                        }

                        //CREATE OR MODIFY VIEWS

                        sql = "CREATE OR REPLACE VIEW MS_PRO_REP_RES";
                        sql += " AS SELECT S.ID,SEC.ID AS SECTID,S.STUDENT_NUMBER,s.Last_Name,s.First_name,C.COURSE_NAME,SG.STANDARDGRADE AS RES FROM standardgradesection sg";
                        sql += " LEFT JOIN STANDARD D ON sg.standardid = D.standardid";
                        sql += " LEFT JOIN STANDARD Di ON sg.standardid = Di.standardid";
                        sql += " LEFT JOIN SECTIONS SEC ON SG.SECTIONSDCID = SEC.dcid";
                        sql += " LEFT JOIN COURSES C ON SEC.COURSE_NUMBER = C.COURSE_NUMBER";
                        sql += " LEFT JOIN TEACHERS T ON SEC.TEACHER = T.id";
                        sql += " LEFT JOIN STUDENTS s ON sg.studentsdcid = s.dcid";
                        sql += " WHERE D.identifier LIKE '%RES%' AND C.COURSE_NAME NOT LIKE '%Explora%' AND C.COURSE_NAME NOT LIKE '%Space%' AND sg.yearid = 28";
                        sql += " AND sg.storecode IN ('Q1', 'S1') AND SG.SCHOOLSDCID = 5 AND S.STUDENT_NUMBER =" + stnumb[a] + "";




                        OracleCommand cmdV1 = new OracleCommand(sql, con);
                        cmdV1.ExecuteNonQuery();

                        sql = "CREATE OR REPLACE VIEW MS_PRO_REP_CON";
                        sql += " AS SELECT S.ID,SEC.ID AS SECID,S.STUDENT_NUMBER,s.Last_Name,s.First_name,C.COURSE_NAME,SG.STANDARDGRADE AS COND FROM standardgradesection sg";
                        sql += " LEFT JOIN STANDARD D ON sg.standardid = D.standardid";
                        sql += " LEFT JOIN SECTIONS SEC ON SG.SECTIONSDCID = SEC.dcid";
                        sql += " LEFT JOIN COURSES C ON SEC.COURSE_NUMBER = C.COURSE_NUMBER";
                        sql += " LEFT JOIN TEACHERS T ON SEC.TEACHER = T.id";
                        sql += " LEFT JOIN STUDENTS s ON sg.studentsdcid = s.dcid";
                        sql += " WHERE D.identifier LIKE '%CON%' AND C.COURSE_NAME NOT LIKE '%Explora%' AND C.COURSE_NAME NOT LIKE '%Advisory%' AND C.COURSE_NAME NOT LIKE '%Space%' AND sg.yearid = 28";
                        sql += " AND sg.storecode IN ('Q1', 'S1') AND SG.SCHOOLSDCID = 5 AND S.STUDENT_NUMBER = " + stnumb[a] + "";


                        OracleCommand cmdV2 = new OracleCommand(sql, con);
                        cmdV2.ExecuteNonQuery();


                        //Advisory Teacher
                        sql = "SELECT DISTINCT C.COURSE_NAME,T.FIRST_NAME||' '||T.LAST_NAME AS TEACHER,R.RES,S.STUDENT_NUMBER,S.FIRST_NAME||' '||S.LAST_NAME AS STUDENT,S.GRADE_LEVEL FROM CC CO";
                        sql += " LEFT JOIN STUDENTS S ON CO.STUDENTID = S.ID";
                        sql += " LEFT JOIN COURSES C ON CO.COURSE_NUMBER = C.COURSE_NUMBER";
                        sql += " LEFT JOIN TEACHERS T ON CO.TEACHERID = T.ID";
                        sql += " LEFT JOIN MS_PRO_REP_RES R ON CO.STUDENTID = R.ID AND CO.SECTIONID = R.SECTID";
                        sql += " WHERE CO.TERMID IN(2800, 2801, 2802)  AND C.COURSE_NAME LIKE '%Advisory%' AND S.STUDENT_NUMBER =" + stnumb[a] + "";

                        OracleCommand cmd = new OracleCommand(sql, con);
                        OracleDataReader odr = cmd.ExecuteReader();
                        while (odr.Read())
                        {
                            Q1AD += odr["COURSE_NAME"].ToString() + '|';
                            Q1AD += odr["TEACHER"].ToString() + '|';
                            Q1AD += odr["RES"].ToString() + '|';
                            Q1AD += odr["STUDENT_NUMBER"].ToString() + '|';
                            Q1AD += odr["STUDENT"].ToString() + '|';
                            Q1AD += odr["GRADE_LEVEL"].ToString() + '|';

                        }

                        sql = "WITH main_query AS(SELECT C.COURSE_NAME, T.FIRST_NAME||' '||T.LAST_NAME AS TEACHER, PG.FINALGRADENAME, PG.GRADE, R.RES, CN.COND,";
                        sql += " TO_CHAR(PG.COMMENT_VALUE) AS COMMENTS, CO.CURRENTABSENCES, CO.CURRENTTARDIES FROM CC CO";
                        sql += " LEFT JOIN STUDENTS S ON CO.STUDENTID = S.ID";
                        sql += " LEFT JOIN COURSES C ON CO.COURSE_NUMBER = C.COURSE_NUMBER";
                        sql += " LEFT JOIN TEACHERS T ON CO.TEACHERID = T.ID";
                        sql += " LEFT JOIN PGFINALGRADES PG ON CO.STUDENTID = PG.STUDENTID AND CO.SECTIONID = PG.SECTIONID";
                        sql += " LEFT JOIN MS_PRO_REP_RES R ON CO.STUDENTID = R.ID AND CO.SECTIONID = R.SECTID";
                        sql += " LEFT JOIN MS_PRO_REP_CON CN ON CO.STUDENTID = CN.ID AND CO.SECTIONID = CN.SECID";
                        sql += " WHERE CO.TERMID IN(2800, 2801, 2802) AND PG.FINALGRADENAME IN('Q1', 'S1') AND C.COURSE_NAME NOT LIKE '%Explora%' AND C.COURSE_NAME NOT LIKE '%Space%'";
                        sql += " AND C.COURSE_NAME NOT LIKE '%Advisory%' AND CO.origsectionid = 0  AND S.STUDENT_NUMBER = " + stnumb[a] + "";
                        sql += " )";
                        sql += " SELECT DISTINCT COURSE_NAME,TEACHER";
                        sql += " ,(SELECT  y.GRADE FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.FINALGRADENAME = 'Q1') SKILL";
                        sql += " ,(SELECT  y.RES FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.FINALGRADENAME = 'Q1') RESP";
                        sql += " ,(SELECT  y.COND FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.FINALGRADENAME = 'Q1') CONDT";
                        sql += " ,(SELECT  y.CURRENTABSENCES FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.FINALGRADENAME = 'Q1') ABS1";
                        sql += " ,(SELECT  y.CURRENTTARDIES FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.FINALGRADENAME = 'Q1') TARDI";
                        sql += " ,(SELECT  y.COMMENTS FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.FINALGRADENAME = 'Q1') Comments";

                        sql += "      FROM main_query M";
                        sql += " ORDER BY COURSE_NAME";
                        OracleCommand cmd1 = new OracleCommand(sql, con);
                        OracleDataReader odr1 = cmd1.ExecuteReader();
                        while (odr1.Read())
                        {
                            Q1DATA += odr1["COURSE_NAME"].ToString() + '|';
                            Q1DATA += odr1["TEACHER"].ToString() + '|';
                            Q1DATA += odr1["SKILL"].ToString() + '|';
                            Q1DATA += odr1["RESP"].ToString() + '|';
                            Q1DATA += odr1["CONDT"].ToString() + '|';
                            Q1DATA += odr1["ABS1"].ToString() + '|';
                            Q1DATA += odr1["TARDI"].ToString() + '|';
                            Q1DATA += odr1["Comments"].ToString() + '^';

                        }
                        if (Q1DATA != "")
                        {
                            // COMMUNITY SERVICE ANAD ST DATA


                            var stTable = Q1DATA.Split('^');
                            var stAdv = Q1AD.Split('|');

                            iTextSharp.text.Image Imagen = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("~/img/WLOGO.jpg"));
                            // Imagen.SetAbsolutePosition(-3, 520);
                            Imagen.ScalePercent(2.5f);


                            PdfPTable HeadT = new PdfPTable(8);
                            HeadT.HorizontalAlignment = Element.ALIGN_CENTER;
                            HeadT.WidthPercentage = 100;

                            PdfPCell logo = new PdfPCell(Imagen);
                            logo.Colspan = 4;
                            logo.Border = 0;
                            logo.HorizontalAlignment = Element.ALIGN_LEFT;
                            logo.Rowspan = 3;
                            logo.Padding = 3;
                            HeadT.AddCell(logo);


                            PdfPCell HS = new PdfPCell(new Phrase("MIDDLE SCHOOL Progress Report", new Font(Font.FontFamily.HELVETICA, 14, Font.BOLD, new BaseColor(135, 0, 27))));
                            HS.HorizontalAlignment = Element.ALIGN_BOTTOM;
                            HS.Colspan = 4;
                            HS.Border = 0;
                            HeadT.AddCell(HS);

                            PdfPCell SQ1 = new PdfPCell(new Phrase("School Year 2018-19 Midsemester 1", new Font(Font.FontFamily.HELVETICA, 14, Font.BOLD, BaseColor.BLACK)));
                            SQ1.HorizontalAlignment = Element.ALIGN_BOTTOM;
                            SQ1.Colspan = 4;
                            SQ1.Border = 0;
                            HeadT.AddCell(SQ1);

                            PdfPCell Pub = new PdfPCell(new Phrase("Published " + DateTime.Now.ToString("MMMM dd, yyyy"), new Font(Font.FontFamily.HELVETICA, 12, Font.ITALIC, BaseColor.BLACK)));
                            Pub.Colspan = 4;
                            Pub.HorizontalAlignment = Element.ALIGN_BOTTOM;
                            Pub.Rowspan = 2;
                            Pub.Border = 0;
                            HeadT.AddCell(Pub);

                            PdfPCell bar1 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                            bar1.HorizontalAlignment = Element.ALIGN_LEFT;
                            bar1.Border = 0;
                            bar1.Colspan = 8;
                            bar1.BackgroundColor = new BaseColor(135, 0, 27);
                            HeadT.AddCell(bar1);

                            PdfPCell stinfo = new PdfPCell(new Phrase("Student Name: " + stAdv[4], new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD, BaseColor.BLACK)));
                            stinfo.HorizontalAlignment = Element.ALIGN_LEFT;
                            stinfo.Border = 0;
                            stinfo.Colspan = 4;
                            stinfo.PaddingTop = 5;
                            HeadT.AddCell(stinfo);

                            PdfPCell messag = new PdfPCell(new Phrase("This report describes progress toward grade level learning expectations, identifies successes and provides guidance for improvement.", new Font(Font.FontFamily.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                            messag.HorizontalAlignment = Element.ALIGN_LEFT;
                            messag.Border = 0;
                            messag.Colspan = 4;
                            messag.Rowspan = 2;
                            HeadT.AddCell(messag);

                            PdfPCell grade = new PdfPCell(new Phrase("Grade: " + stAdv[5], new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD, BaseColor.BLACK)));
                            grade.HorizontalAlignment = Element.ALIGN_LEFT;
                            grade.Border = 0;
                            grade.Colspan = 2;
                            HeadT.AddCell(grade);
                            PdfPCell stid = new PdfPCell(new Phrase("StudentID: " + stAdv[3], new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD, BaseColor.WHITE)));
                            stid.HorizontalAlignment = Element.ALIGN_LEFT;
                            stid.Border = 0;
                            stid.Colspan = 2;
                            stid.PaddingBottom = 5;
                            HeadT.AddCell(stid);

                            PdfPCell boh = new PdfPCell(new Phrase("Advisory: " + stAdv[1] + " (R:" + stAdv[2] + ")", new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD, BaseColor.BLACK)));
                            boh.HorizontalAlignment = Element.ALIGN_LEFT;
                            boh.Border = 0;
                            boh.Colspan = 4;
                            HeadT.AddCell(boh);

                            PdfPCell legdS = new PdfPCell(new Phrase("S = Subject Area Skills       A = Absences" + Environment.NewLine + "R = Responsibility               T = Tardies" + Environment.NewLine + "C = Conduct                        ", new Font(Font.FontFamily.HELVETICA, 9, Font.NORMAL, BaseColor.BLACK)));
                            legdS.HorizontalAlignment = Element.ALIGN_LEFT;
                            legdS.Border = 0;
                            legdS.Colspan = 3;
                            legdS.Rowspan = 2;
                            legdS.PaddingBottom = 10;
                            HeadT.AddCell(legdS);

                            PdfPCell LINK = new PdfPCell();
                            var c = new Chunk("Standards Score" + Environment.NewLine + "Assessment Key", new Font(Font.FontFamily.HELVETICA, 9, Font.UNDERLINE, BaseColor.BLUE));
                            c.SetAnchor("http://bit.ly/MSStdAssessKey");
                            LINK.AddElement(c);
                            LINK.Border = 0;
                            LINK.Rowspan = 2;
                            HeadT.AddCell(LINK);

                            PdfPTable GradeTable = new PdfPTable(18);
                            GradeTable.HorizontalAlignment = Element.ALIGN_CENTER;
                            GradeTable.WidthPercentage = 100;

                            PdfPCell Course = new PdfPCell(new Phrase("Course", new Font(Font.FontFamily.HELVETICA, 12.0F, Font.BOLD, BaseColor.WHITE)));
                            Course.HorizontalAlignment = Element.ALIGN_LEFT;

                            Course.BackgroundColor = new BaseColor(135, 0, 27);
                            Course.BorderWidth = 1F;
                            Course.Colspan = 7;
                            Course.PaddingBottom = 5;
                            GradeTable.AddCell(Course);
                            PdfPCell Teacher = new PdfPCell(new Phrase("Teacher", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                            Teacher.HorizontalAlignment = Element.ALIGN_LEFT;

                            Teacher.BackgroundColor = new BaseColor(135, 0, 27);
                            Teacher.BorderWidth = 1F;
                            Teacher.Colspan = 5;
                            Teacher.PaddingBottom = 5;
                            GradeTable.AddCell(Teacher);
                            PdfPCell Q1 = new PdfPCell(new Phrase("S", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                            Q1.HorizontalAlignment = Element.ALIGN_CENTER;
                            Q1.BackgroundColor = new BaseColor(135, 0, 27);
                            Q1.BorderWidth = 1F;
                            Q1.PaddingBottom = 5;
                            Q1.Colspan = 2;
                            GradeTable.AddCell(Q1);
                            PdfPCell Res = new PdfPCell(new Phrase("R", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                            Res.HorizontalAlignment = Element.ALIGN_CENTER;
                            Res.BackgroundColor = new BaseColor(135, 0, 27);
                            Res.BorderWidth = 1F;
                            Res.PaddingBottom = 5;
                            GradeTable.AddCell(Res);
                            PdfPCell Cond = new PdfPCell(new Phrase("C", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                            Cond.HorizontalAlignment = Element.ALIGN_CENTER;
                            Cond.BackgroundColor = new BaseColor(135, 0, 27);
                            Cond.BorderWidth = 1F;
                            Cond.PaddingBottom = 5;
                            GradeTable.AddCell(Cond);
                            PdfPCell ABS = new PdfPCell(new Phrase("A/T", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                            ABS.HorizontalAlignment = Element.ALIGN_CENTER;
                            ABS.BackgroundColor = new BaseColor(135, 0, 27);
                            ABS.BorderWidth = 1F;
                            ABS.Colspan = 2;
                            ABS.PaddingBottom = 5;
                            GradeTable.AddCell(ABS);



                            PdfPCell CO;
                            PdfPCell TE;
                            PdfPCell QU1;
                            PdfPCell COM1;
                            PdfPCell ABS1;
                            PdfPCell R1;
                            PdfPCell Co1;

                            for (int i = 0; i < stTable.Length - 1; i++)
                            {
                                var nfila = stTable[i].Split('|');

                                CO = new PdfPCell(new Phrase(nfila[0], new Font(Font.FontFamily.HELVETICA, 11F, Font.NORMAL, BaseColor.BLACK)));
                                CO.HorizontalAlignment = Element.ALIGN_LEFT;
                                CO.BorderWidth = 0.5F;
                                CO.Colspan = 7;
                                CO.PaddingBottom = 3;
                                CO.BackgroundColor = new BaseColor(235, 235, 235);
                                CO.BorderColor = BaseColor.GRAY;
                                GradeTable.AddCell(CO);

                                TE = new PdfPCell(new Phrase(nfila[1], new Font(Font.FontFamily.HELVETICA, 11F, Font.NORMAL, BaseColor.BLACK)));
                                TE.HorizontalAlignment = Element.ALIGN_LEFT;
                                TE.BorderWidth = 0.5F;
                                TE.Colspan = 5;
                                TE.PaddingBottom = 3;
                                TE.BackgroundColor = new BaseColor(235, 235, 235);
                                TE.BorderColor = BaseColor.GRAY;
                                GradeTable.AddCell(TE);

                                QU1 = new PdfPCell(new Phrase(nfila[2], new Font(Font.FontFamily.HELVETICA, 11F, Font.NORMAL, BaseColor.BLACK)));
                                QU1.HorizontalAlignment = Element.ALIGN_CENTER;
                                QU1.BorderWidth = 0.5F;
                                QU1.Colspan = 2;
                                QU1.PaddingBottom = 3;
                                QU1.BackgroundColor = new BaseColor(235, 235, 235);
                                QU1.BorderColor = BaseColor.GRAY;
                                GradeTable.AddCell(QU1);

                                R1 = new PdfPCell(new Phrase(nfila[3], new Font(Font.FontFamily.HELVETICA, 11F, Font.NORMAL, BaseColor.BLACK)));
                                R1.HorizontalAlignment = Element.ALIGN_CENTER;
                                R1.BorderWidth = 0.5F;
                                R1.PaddingBottom = 3;
                                R1.BackgroundColor = new BaseColor(235, 235, 235);
                                R1.BorderColor = BaseColor.GRAY;
                                GradeTable.AddCell(R1);

                                Co1 = new PdfPCell(new Phrase(nfila[4], new Font(Font.FontFamily.HELVETICA, 11F, Font.NORMAL, BaseColor.BLACK)));
                                Co1.HorizontalAlignment = Element.ALIGN_CENTER;
                                Co1.BorderWidth = 0.5F;
                                Co1.PaddingBottom = 3;
                                Co1.BackgroundColor = new BaseColor(235, 235, 235);
                                Co1.BorderColor = BaseColor.GRAY;
                                GradeTable.AddCell(Co1);

                                ABS1 = new PdfPCell(new Phrase(nfila[5] + '/' + nfila[6], new Font(Font.FontFamily.HELVETICA, 11F, Font.NORMAL, BaseColor.BLACK)));
                                ABS1.HorizontalAlignment = Element.ALIGN_CENTER;
                                ABS1.BorderWidth = 0.5F;
                                ABS1.Colspan = 2;
                                ABS1.PaddingBottom = 3;
                                ABS1.BackgroundColor = new BaseColor(235, 235, 235);
                                ABS1.BorderColor = BaseColor.GRAY;
                                GradeTable.AddCell(ABS1);

                                COM1 = new PdfPCell(new Phrase(nfila[7], new Font(Font.FontFamily.HELVETICA, 11F, Font.NORMAL, BaseColor.BLACK)));
                                COM1.HorizontalAlignment = Element.ALIGN_LEFT;
                                COM1.BorderWidth = 0;
                                COM1.Colspan = 18;
                                COM1.PaddingBottom = 8;
                                //COM1.BorderColor = BaseColor.LIGHT_GRAY;
                                GradeTable.AddCell(COM1);


                        }

                            //PdfPTable FOOT = new PdfPTable(8);
                            //FOOT.HorizontalAlignment = Element.ALIGN_LEFT;
                            //FOOT.WidthPercentage = 100;

                            //PdfPCell spa3 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.WHITE)));
                            //spa3.Colspan = 8;
                            //spa3.Border = 0;
                            //FOOT.AddCell(spa3);

                            
                         

                            //documento.Add(stfoto);
                            documento.Add(HeadT);
                            documento.Add(GradeTable);
                            //documento.Add(FOOT);

                            //Process prc = new System.Diagnostics.Process();
                            //prc.StartInfo.FileName = fileName;
                            //prc.Start();
                        }
                        else
                        {
                            con.Close();

                        }
                        documento.NewPage();
                    }

                    con.Close();

                }
                else
                {
                    if (con.State == ConnectionState.Closed)
                    {
                        con.Open();
                    }
                    fname = "MS_ProgressReport_" + DateTime.Now.DayOfYear + DateTime.Now.Month + DateTime.Now.Day + DateTime.Now.Millisecond + ".pdf";
                    fileName = HttpContext.Current.Server.MapPath("~/RepoFiles/" + fname);
                    PdfWriter.GetInstance(documento, new FileStream(fileName, FileMode.Create));
                    documento.Open();

                        string Q1DATA = string.Empty;
                        string Q1AD = string.Empty;

                        if (con.State == ConnectionState.Closed)
                        {
                            con.Open();
                        }

                    //CREATE OR MODIFY VIEWS

                    sql = "CREATE OR REPLACE VIEW MS_PRO_REP_RES";
                    sql += " AS SELECT S.ID,SEC.ID AS SECTID,S.STUDENT_NUMBER,s.Last_Name,s.First_name,C.COURSE_NAME,SG.STANDARDGRADE AS RES FROM standardgradesection sg";
                    sql += " LEFT JOIN STANDARD D ON sg.standardid = D.standardid";
                    sql += " LEFT JOIN STANDARD Di ON sg.standardid = Di.standardid";
                    sql += " LEFT JOIN SECTIONS SEC ON SG.SECTIONSDCID = SEC.dcid";
                    sql += " LEFT JOIN COURSES C ON SEC.COURSE_NUMBER = C.COURSE_NUMBER";
                    sql += " LEFT JOIN TEACHERS T ON SEC.TEACHER = T.id";
                    sql += " LEFT JOIN STUDENTS s ON sg.studentsdcid = s.dcid";
                    sql += " WHERE D.identifier LIKE '%RES%' AND C.COURSE_NAME NOT LIKE '%Explora%' AND C.COURSE_NAME NOT LIKE '%Space%' AND sg.yearid = 28";
                    sql += " AND sg.storecode IN ('Q1', 'S1') AND SG.SCHOOLSDCID = 5 AND S.STUDENT_NUMBER =" + stnum + "";




                    OracleCommand cmdV1 = new OracleCommand(sql, con);
                        cmdV1.ExecuteNonQuery();

                    sql = "CREATE OR REPLACE VIEW MS_PRO_REP_CON";
                    sql += " AS SELECT S.ID,SEC.ID AS SECID,S.STUDENT_NUMBER,s.Last_Name,s.First_name,C.COURSE_NAME,SG.STANDARDGRADE AS COND FROM standardgradesection sg";
                    sql += " LEFT JOIN STANDARD D ON sg.standardid = D.standardid";
                    sql += " LEFT JOIN SECTIONS SEC ON SG.SECTIONSDCID = SEC.dcid";
                    sql += " LEFT JOIN COURSES C ON SEC.COURSE_NUMBER = C.COURSE_NUMBER";
                    sql += " LEFT JOIN TEACHERS T ON SEC.TEACHER = T.id";
                    sql += " LEFT JOIN STUDENTS s ON sg.studentsdcid = s.dcid";
                    sql += " WHERE D.identifier LIKE '%CON%' AND C.COURSE_NAME NOT LIKE '%Explora%' AND C.COURSE_NAME NOT LIKE '%Advisory%' AND C.COURSE_NAME NOT LIKE '%Space%' AND sg.yearid = 28";
                    sql += " AND sg.storecode IN ('Q1', 'S1') AND SG.SCHOOLSDCID = 5 AND S.STUDENT_NUMBER = " + stnum + "";


                        OracleCommand cmdV2 = new OracleCommand(sql, con);
                        cmdV2.ExecuteNonQuery();


                        //Advisory Teacher
                        sql = "SELECT DISTINCT C.COURSE_NAME,T.FIRST_NAME||' '||T.LAST_NAME AS TEACHER,R.RES,S.STUDENT_NUMBER,S.FIRST_NAME||' '||S.LAST_NAME AS STUDENT,S.GRADE_LEVEL FROM CC CO";
                        sql += " LEFT JOIN STUDENTS S ON CO.STUDENTID = S.ID";
                        sql += " LEFT JOIN COURSES C ON CO.COURSE_NUMBER = C.COURSE_NUMBER";
                        sql += " LEFT JOIN TEACHERS T ON CO.TEACHERID = T.ID";
                        sql += " LEFT JOIN MS_PRO_REP_RES R ON CO.STUDENTID = R.ID AND CO.SECTIONID = R.SECTID";
                        sql += " WHERE CO.TERMID IN(2800, 2801, 2802)  AND C.COURSE_NAME LIKE '%Advisory%' AND S.STUDENT_NUMBER =" + stnum + "";

                        OracleCommand cmd = new OracleCommand(sql, con);
                        OracleDataReader odr = cmd.ExecuteReader();
                        while (odr.Read())
                        {
                            Q1AD += odr["COURSE_NAME"].ToString() + '|';
                            Q1AD += odr["TEACHER"].ToString() + '|';
                            Q1AD += odr["RES"].ToString() + '|';
                            Q1AD += odr["STUDENT_NUMBER"].ToString() + '|';
                            Q1AD += odr["STUDENT"].ToString() + '|';
                            Q1AD += odr["GRADE_LEVEL"].ToString() + '|';

                        }

                    sql = "WITH main_query AS(SELECT C.COURSE_NAME, T.FIRST_NAME||' '||T.LAST_NAME AS TEACHER, PG.FINALGRADENAME, PG.GRADE, R.RES, CN.COND,";
                      sql += " TO_CHAR(PG.COMMENT_VALUE) AS COMMENTS, CO.CURRENTABSENCES, CO.CURRENTTARDIES FROM CC CO";
                       sql += " LEFT JOIN STUDENTS S ON CO.STUDENTID = S.ID";
                       sql += " LEFT JOIN COURSES C ON CO.COURSE_NUMBER = C.COURSE_NUMBER";
                       sql += " LEFT JOIN TEACHERS T ON CO.TEACHERID = T.ID";
                       sql += " LEFT JOIN PGFINALGRADES PG ON CO.STUDENTID = PG.STUDENTID AND CO.SECTIONID = PG.SECTIONID";
                       sql += " LEFT JOIN MS_PRO_REP_RES R ON CO.STUDENTID = R.ID AND CO.SECTIONID = R.SECTID";
                       sql += " LEFT JOIN MS_PRO_REP_CON CN ON CO.STUDENTID = CN.ID AND CO.SECTIONID = CN.SECID";
                       sql += " WHERE CO.TERMID IN(2800, 2801, 2802) AND PG.FINALGRADENAME IN('Q1', 'S1') AND C.COURSE_NAME NOT LIKE '%Explora%' AND C.COURSE_NAME NOT LIKE '%Space%'";
                       sql += " AND C.COURSE_NAME NOT LIKE '%Advisory%' AND CO.origsectionid = 0  AND S.STUDENT_NUMBER = " + stnum + ""; 
                         sql += " )";
                        sql += " SELECT DISTINCT COURSE_NAME,TEACHER";
                         sql += " ,(SELECT  y.GRADE FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.FINALGRADENAME = 'Q1') SKILL";
                    sql += " ,(SELECT  y.RES FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.FINALGRADENAME = 'Q1') RESP";
                     sql += " ,(SELECT  y.COND FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.FINALGRADENAME = 'Q1') CONDT";
                    sql += " ,(SELECT  y.CURRENTABSENCES FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.FINALGRADENAME = 'Q1') ABS1";
                     sql += " ,(SELECT  y.CURRENTTARDIES FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.FINALGRADENAME = 'Q1') TARDI";
                    sql += " ,(SELECT  y.COMMENTS FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.FINALGRADENAME = 'Q1') Comments";

                    sql += "      FROM main_query M";
                     sql += " ORDER BY COURSE_NAME";


                    OracleCommand cmd1 = new OracleCommand(sql, con);
                        OracleDataReader odr1 = cmd1.ExecuteReader();
                        while (odr1.Read())
                        {
                            Q1DATA += odr1["COURSE_NAME"].ToString() + '|';
                            Q1DATA += odr1["TEACHER"].ToString() + '|';
                            Q1DATA += odr1["SKILL"].ToString() + '|';
                            Q1DATA += odr1["RESP"].ToString() + '|';
                            Q1DATA += odr1["CONDT"].ToString() + '|';
                            Q1DATA += odr1["ABS1"].ToString() + '|';
                            Q1DATA += odr1["TARDI"].ToString() + '|';
                            Q1DATA += odr1["Comments"].ToString() + '^';

                        }
                        if (Q1DATA != "")
                        {
                            // COMMUNITY SERVICE ANAD ST DATA


                            var stTable = Q1DATA.Split('^');
                            var stAdv = Q1AD.Split('|');

                            iTextSharp.text.Image Imagen = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("~/img/WLOGO.jpg"));
                            // Imagen.SetAbsolutePosition(-3, 520);
                            Imagen.ScalePercent(2.5f);


                            PdfPTable HeadT = new PdfPTable(8);
                            HeadT.HorizontalAlignment = Element.ALIGN_CENTER;
                            HeadT.WidthPercentage = 100;

                            PdfPCell logo = new PdfPCell(Imagen);
                            logo.Colspan = 4;
                            logo.Border = 0;
                            logo.HorizontalAlignment = Element.ALIGN_LEFT;
                            logo.Rowspan = 3;
                            logo.Padding = 3;
                            HeadT.AddCell(logo);


                        PdfPCell HS = new PdfPCell(new Phrase("MIDDLE SCHOOL Progress Report", new Font(Font.FontFamily.HELVETICA, 14, Font.BOLD, new BaseColor(135, 0, 27))));
                        HS.HorizontalAlignment = Element.ALIGN_BOTTOM;
                        HS.Colspan = 4;
                        HS.Border = 0;
                        HeadT.AddCell(HS);

                        PdfPCell SQ1 = new PdfPCell(new Phrase("School Year 2018-19 Midsemester 1", new Font(Font.FontFamily.HELVETICA, 14, Font.BOLD, BaseColor.BLACK)));
                        SQ1.HorizontalAlignment = Element.ALIGN_BOTTOM;
                        SQ1.Colspan = 4;
                        SQ1.Border = 0;
                        HeadT.AddCell(SQ1);

                        PdfPCell Pub = new PdfPCell(new Phrase("Published " + DateTime.Now.ToString("MMMM dd, yyyy"), new Font(Font.FontFamily.HELVETICA, 12, Font.ITALIC, BaseColor.BLACK)));
                        Pub.Colspan = 4;
                        Pub.HorizontalAlignment = Element.ALIGN_BOTTOM;
                        Pub.Rowspan = 2;
                        Pub.Border = 0;
                        HeadT.AddCell(Pub);

                        PdfPCell bar1 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                        bar1.HorizontalAlignment = Element.ALIGN_LEFT;
                        bar1.Border = 0;
                        bar1.Colspan = 8;
                        bar1.BackgroundColor = new BaseColor(135, 0, 27);
                        HeadT.AddCell(bar1);

                        PdfPCell stinfo = new PdfPCell(new Phrase("Student Name: " + stAdv[4], new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD, BaseColor.BLACK)));
                        stinfo.HorizontalAlignment = Element.ALIGN_LEFT;
                        stinfo.Border = 0;
                        stinfo.Colspan = 4;
                        stinfo.PaddingTop = 5;
                        HeadT.AddCell(stinfo);

                        PdfPCell messag = new PdfPCell(new Phrase("This report describes progress toward grade level learning expectations, identifies successes and provides guidance for improvement.", new Font(Font.FontFamily.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                        messag.HorizontalAlignment = Element.ALIGN_LEFT;
                        messag.Border = 0;
                        messag.Colspan = 4;
                        messag.Rowspan = 2;
                        HeadT.AddCell(messag);

                        PdfPCell grade = new PdfPCell(new Phrase("Grade: " + stAdv[5], new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD, BaseColor.BLACK)));
                        grade.HorizontalAlignment = Element.ALIGN_LEFT;
                        grade.Border = 0;
                        grade.Colspan = 2;
                        HeadT.AddCell(grade);
                        PdfPCell stid = new PdfPCell(new Phrase("StudentID: " + stAdv[3], new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD, BaseColor.WHITE)));
                        stid.HorizontalAlignment = Element.ALIGN_LEFT;
                        stid.Border = 0;
                        stid.Colspan = 2;
                        stid.PaddingBottom = 5;
                        HeadT.AddCell(stid);

                        PdfPCell boh = new PdfPCell(new Phrase("Advisory: " + stAdv[1] + " (R:" + stAdv[2] + ")", new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD, BaseColor.BLACK)));
                        boh.HorizontalAlignment = Element.ALIGN_LEFT;
                        boh.Border = 0;
                        boh.Colspan = 4;
                        HeadT.AddCell(boh);

                        PdfPCell legdS = new PdfPCell(new Phrase("S = Subject Area Skills       A = Absences" + Environment.NewLine + "R = Responsibility               T = Tardies" + Environment.NewLine + "C = Conduct                        ", new Font(Font.FontFamily.HELVETICA, 9, Font.NORMAL, BaseColor.BLACK)));
                        legdS.HorizontalAlignment = Element.ALIGN_LEFT;
                        legdS.Border = 0;
                        legdS.Colspan = 3;
                        legdS.Rowspan = 2;
                        legdS.PaddingBottom = 10;
                        HeadT.AddCell(legdS);

                        PdfPCell LINK = new PdfPCell();
                        var c = new Chunk("Standards Score" + Environment.NewLine + "Assessment Key", new Font(Font.FontFamily.HELVETICA, 9, Font.UNDERLINE, BaseColor.BLUE));
                        c.SetAnchor("http://bit.ly/MSStdAssessKey");
                        LINK.AddElement(c);
                        LINK.Border = 0;
                        LINK.Rowspan = 2;
                        HeadT.AddCell(LINK);

                        PdfPTable GradeTable = new PdfPTable(18);
                        GradeTable.HorizontalAlignment = Element.ALIGN_CENTER;
                        GradeTable.WidthPercentage = 100;

                        PdfPCell Course = new PdfPCell(new Phrase("Course", new Font(Font.FontFamily.HELVETICA, 12.0F, Font.BOLD, BaseColor.WHITE)));
                        Course.HorizontalAlignment = Element.ALIGN_LEFT;

                        Course.BackgroundColor = new BaseColor(135, 0, 27);
                        Course.BorderWidth = 1F;
                        Course.Colspan = 7;
                        Course.PaddingBottom = 5;
                        GradeTable.AddCell(Course);
                        PdfPCell Teacher = new PdfPCell(new Phrase("Teacher", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                        Teacher.HorizontalAlignment = Element.ALIGN_LEFT;

                        Teacher.BackgroundColor = new BaseColor(135, 0, 27);
                        Teacher.BorderWidth = 1F;
                        Teacher.Colspan = 5;
                        Teacher.PaddingBottom = 5;
                        GradeTable.AddCell(Teacher);
                        PdfPCell Q1 = new PdfPCell(new Phrase("S", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                        Q1.HorizontalAlignment = Element.ALIGN_CENTER;
                        Q1.BackgroundColor = new BaseColor(135, 0, 27);
                        Q1.BorderWidth = 1F;
                        Q1.PaddingBottom = 5;
                        Q1.Colspan = 2;
                        GradeTable.AddCell(Q1);
                        PdfPCell Res = new PdfPCell(new Phrase("R", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                        Res.HorizontalAlignment = Element.ALIGN_CENTER;
                        Res.BackgroundColor = new BaseColor(135, 0, 27);
                        Res.BorderWidth = 1F;
                        Res.PaddingBottom = 5;
                        GradeTable.AddCell(Res);
                        PdfPCell Cond = new PdfPCell(new Phrase("C", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                        Cond.HorizontalAlignment = Element.ALIGN_CENTER;
                        Cond.BackgroundColor = new BaseColor(135, 0, 27);
                        Cond.BorderWidth = 1F;
                        Cond.PaddingBottom = 5;
                        GradeTable.AddCell(Cond);
                        PdfPCell ABS = new PdfPCell(new Phrase("A/T", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                        ABS.HorizontalAlignment = Element.ALIGN_CENTER;
                        ABS.BackgroundColor = new BaseColor(135, 0, 27);
                        ABS.BorderWidth = 1F;
                        ABS.Colspan = 2;
                        ABS.PaddingBottom = 5;
                        GradeTable.AddCell(ABS);



                        PdfPCell CO;
                        PdfPCell TE;
                        PdfPCell QU1;
                        PdfPCell COM1;
                        PdfPCell ABS1;
                        PdfPCell R1;
                        PdfPCell Co1;

                        for (int i = 0; i < stTable.Length - 1; i++)
                        {
                            var nfila = stTable[i].Split('|');

                            CO = new PdfPCell(new Phrase(nfila[0], new Font(Font.FontFamily.HELVETICA, 11F, Font.NORMAL, BaseColor.BLACK)));
                            CO.HorizontalAlignment = Element.ALIGN_LEFT;
                            CO.BorderWidth = 0.5F;
                            CO.Colspan = 7;
                            CO.PaddingBottom = 3;
                            CO.BackgroundColor = new BaseColor(235, 235, 235);
                            CO.BorderColor = BaseColor.GRAY;
                            GradeTable.AddCell(CO);

                            TE = new PdfPCell(new Phrase(nfila[1], new Font(Font.FontFamily.HELVETICA, 11F, Font.NORMAL, BaseColor.BLACK)));
                            TE.HorizontalAlignment = Element.ALIGN_LEFT;
                            TE.BorderWidth = 0.5F;
                            TE.Colspan = 5;
                            TE.PaddingBottom = 3;
                            TE.BackgroundColor = new BaseColor(235, 235, 235);
                            TE.BorderColor = BaseColor.GRAY;
                            GradeTable.AddCell(TE);

                            QU1 = new PdfPCell(new Phrase(nfila[2], new Font(Font.FontFamily.HELVETICA, 11F, Font.NORMAL, BaseColor.BLACK)));
                            QU1.HorizontalAlignment = Element.ALIGN_CENTER;
                            QU1.BorderWidth = 0.5F;
                            QU1.Colspan = 2;
                            QU1.PaddingBottom = 3;
                            QU1.BackgroundColor = new BaseColor(235, 235, 235);
                            QU1.BorderColor = BaseColor.GRAY;
                            GradeTable.AddCell(QU1);

                            R1 = new PdfPCell(new Phrase(nfila[3], new Font(Font.FontFamily.HELVETICA, 11F, Font.NORMAL, BaseColor.BLACK)));
                            R1.HorizontalAlignment = Element.ALIGN_CENTER;
                            R1.BorderWidth = 0.5F;
                            R1.PaddingBottom = 3;
                            R1.BackgroundColor = new BaseColor(235, 235, 235);
                            R1.BorderColor = BaseColor.GRAY;
                            GradeTable.AddCell(R1);

                            Co1 = new PdfPCell(new Phrase(nfila[4], new Font(Font.FontFamily.HELVETICA, 11F, Font.NORMAL, BaseColor.BLACK)));
                            Co1.HorizontalAlignment = Element.ALIGN_CENTER;
                            Co1.BorderWidth = 0.5F;
                            Co1.PaddingBottom = 3;
                            Co1.BackgroundColor = new BaseColor(235, 235, 235);
                            Co1.BorderColor = BaseColor.GRAY;
                            GradeTable.AddCell(Co1);

                            ABS1 = new PdfPCell(new Phrase(nfila[5] + '/' + nfila[6], new Font(Font.FontFamily.HELVETICA, 11F, Font.NORMAL, BaseColor.BLACK)));
                            ABS1.HorizontalAlignment = Element.ALIGN_CENTER;
                            ABS1.BorderWidth = 0.5F;
                            ABS1.Colspan = 2;
                            ABS1.PaddingBottom = 3;
                            ABS1.BackgroundColor = new BaseColor(235, 235, 235);
                            ABS1.BorderColor = BaseColor.GRAY;
                            GradeTable.AddCell(ABS1);

                            COM1 = new PdfPCell(new Phrase(nfila[7], new Font(Font.FontFamily.HELVETICA, 11F, Font.NORMAL, BaseColor.BLACK)));
                            COM1.HorizontalAlignment = Element.ALIGN_LEFT;
                            COM1.BorderWidth = 0;
                            COM1.Colspan = 18;
                            COM1.PaddingBottom = 8;
                            //COM1.BorderColor = BaseColor.LIGHT_GRAY;
                            GradeTable.AddCell(COM1);


                        }

                        


                        //documento.Add(stfoto);
                        documento.Add(HeadT);
                            documento.Add(GradeTable);
                           // documento.Add(FOOT);



                        //Process prc = new System.Diagnostics.Process();
                        //prc.StartInfo.FileName = fileName;
                        //prc.Start();
                    }
                    else
                    {
                        con.Close();
                        fname = "";
                    }
                    con.Close();

                }

                documento.Close();
                con.Dispose();
            }
            catch (Exception ex)
            {
                throw;
            }
            return fname;
        }

        [WebMethod]
        public static string HSQ1(string stnum)
        {
            string sql = string.Empty;
            // int stnum = 4388;


            string fname = string.Empty;
            string fileName = string.Empty;
            OracleConnection con = new OracleConnection();
            con.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings["conn"].ConnectionString;
            Document documento = new Document(PageSize.LETTER, 10, 10, 5, 5);
            try
            {
                
                if (stnum.IndexOf(';') > -1)
                {
                    var stnumb = stnum.Split(';');
                    fname = "HS_ProgressReport_" + DateTime.Now.DayOfYear + DateTime.Now.Month + DateTime.Now.Day + DateTime.Now.Millisecond + ".pdf";
                    fileName = HttpContext.Current.Server.MapPath("~/RepoFiles/" + fname);
                    PdfWriter.GetInstance(documento, new FileStream(fileName, FileMode.Create));
                    documento.Open();

                    for (int a = 0; a < stnumb.Length; a++)
                    {
                        string stdata = string.Empty;
                        string S1GPA = string.Empty;
                        string CUGPA = string.Empty;
                        string datat = string.Empty;
                        if (con.State == ConnectionState.Closed)
                        {
                            con.Open();
                        }
                        //GRADES VALUES
                        sql = "WITH main_query AS ( SELECT C.COURSE_NAME,T.FIRST_NAME||' '||T.LAST_NAME AS TEACHER, PG.FINALGRADENAME,PG.GRADE,";
                        sql += " TO_CHAR(PG.COMMENT_VALUE) AS COMMENTS,CO.CURRENTABSENCES,CO.CURRENTTARDIES FROM CC CO";
                        sql += " LEFT JOIN STUDENTS S ON CO.STUDENTID=S.ID";
                        sql += " LEFT JOIN COURSES C ON CO.COURSE_NUMBER=C.COURSE_NUMBER";
                        sql += " LEFT JOIN TEACHERS T ON CO.TEACHERID=T.ID";
                        sql += " LEFT JOIN PGFINALGRADES PG ON CO.SECTIONID=PG.SECTIONID AND CO.STUDENTID=PG.STUDENTID";
                        sql += " WHERE CO.TERMID IN(2800,2801,2802) AND C.COURSE_NAME NOT LIKE '%Bohio%' AND PG.FINALGRADENAME IN ('Q1','S1') AND S.STUDENT_NUMBER=" + stnumb[a] + "";
                        sql += " )";
                        sql += " SELECT DISTINCT COURSE_NAME,TEACHER";
                        sql += " ,(SELECT  y.GRADE FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.FINALGRADENAME='Q1') Q1";
                        sql += " ,(SELECT  y.COMMENTS FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.FINALGRADENAME='Q1') Comments";
                        sql += " ,(SELECT  y.CURRENTABSENCES FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.FINALGRADENAME='Q1') ABS1";
                        sql += " ,(SELECT  y.CURRENTTARDIES FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.FINALGRADENAME='Q1') TARDI";
                        sql += "  FROM	main_query M";
                        sql += " ORDER BY COURSE_NAME";

                        OracleCommand cmd = new OracleCommand(sql, con);
                        OracleDataReader odr = cmd.ExecuteReader();
                        while (odr.Read())
                        {
                            datat += odr["COURSE_NAME"].ToString() + '|';
                            datat += odr["TEACHER"].ToString() + '|';
                            datat += odr["Q1"].ToString() + '|';
                            datat += odr["Comments"].ToString() + '|';
                            datat += odr["ABS1"].ToString() + '|';
                            datat += odr["TARDI"].ToString() + '^';

                        }
                        if (datat != "")
                        {
                            // COMMUNITY SERVICE ANAD ST DATA

                            sql = "SELECT S.STUDENT_NUMBER,S.FIRST_NAME||' '||S.LAST_NAME AS STUDENT,S.GRADE_LEVEL,C.COURSE_NAME,T.FIRST_NAME||' '||T.LAST_NAME AS TEACHER,CO.HS_SERVICE_HOURS_CURRENT,CO.HS_SERVICE_HOURS_CURRENT_IN,";
                            sql += " CO.HS_SERVICE_HOURS_CURRENT_OUTRE,CO.HS_TOTAL_SERVICE_HOURS  FROM CC CO";
                            sql += " LEFT JOIN STUDENTS S ON CO.STUDENTID=S.ID";
                            sql += " LEFT JOIN COURSES C ON CO.COURSE_NUMBER=C.COURSE_NUMBER";
                            sql += " LEFT JOIN TEACHERS T ON CO.TEACHERID=T.ID";
                            sql += " LEFT JOIN U_COMMUNITYSERVICE CO ON S.DCID=CO.STUDENTSDCID";
                            sql += " WHERE CO.TERMID IN(2800) AND COURSE_NAME LIKE '%Bohio%' AND S.STUDENT_NUMBER=" + stnumb[a] + "";

                            OracleCommand cmd1 = new OracleCommand(sql, con);
                            OracleDataReader odr1 = cmd1.ExecuteReader();
                            while (odr1.Read())
                            {
                                stdata += odr1["STUDENT_NUMBER"].ToString() + '|';
                                stdata += odr1["STUDENT"].ToString() + '|';
                                stdata += odr1["GRADE_LEVEL"].ToString() + '|';
                                stdata += odr1["COURSE_NAME"].ToString() + '|';
                                stdata += odr1["TEACHER"].ToString() + '|';
                                stdata += odr1["HS_SERVICE_HOURS_CURRENT"].ToString() + '|';
                                stdata += odr1["HS_SERVICE_HOURS_CURRENT_IN"].ToString() + '|';
                                stdata += odr1["HS_SERVICE_HOURS_CURRENT_OUTRE"].ToString() + '|';
                                stdata += odr1["HS_TOTAL_SERVICE_HOURS"].ToString() + '|';

                            }

                            var stTable = datat.Split('^');
                            var stcomm = stdata.Split('|');
                            
                            iTextSharp.text.Image Imagen = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("~/img/WLOGO.jpg"));
                           // Imagen.SetAbsolutePosition(-3, 520);
                            Imagen.ScalePercent(2.5f);


                            PdfPTable HeadT = new PdfPTable(8);
                            HeadT.HorizontalAlignment = Element.ALIGN_CENTER;
                            HeadT.WidthPercentage = 100;

                            PdfPCell logo = new PdfPCell(Imagen);
                            logo.Colspan = 4;
                            logo.Border = 0;
                            logo.HorizontalAlignment = Element.ALIGN_LEFT;
                            logo.Rowspan = 3;
                            logo.Padding = 3;
                            HeadT.AddCell(logo);


                            PdfPCell HS = new PdfPCell(new Phrase("HIGH SCHOOL Progress Report", new Font(Font.FontFamily.HELVETICA, 14, Font.BOLD, new BaseColor(135, 0, 27))));
                            HS.HorizontalAlignment = Element.ALIGN_LEFT;
                            HS.Colspan = 4;
                            HS.Border = 0;
                            HeadT.AddCell(HS);

                            PdfPCell SQ1 = new PdfPCell(new Phrase("School Year 2018-19 Midsemester 1 (Q1)", new Font(Font.FontFamily.HELVETICA, 14, Font.BOLD, BaseColor.BLACK)));
                            SQ1.HorizontalAlignment = Element.ALIGN_LEFT;
                            SQ1.Colspan = 4;
                            SQ1.Border = 0;
                            HeadT.AddCell(SQ1);

                            PdfPCell Pub = new PdfPCell(new Phrase("Published " + DateTime.Now.ToString("MMMM dd, yyyy"), new Font(Font.FontFamily.HELVETICA, 12, Font.ITALIC, BaseColor.BLACK)));
                            Pub.Colspan = 4;
                            Pub.HorizontalAlignment = Element.ALIGN_LEFT;
                            Pub.Border = 0;
                            HeadT.AddCell(Pub);

                            PdfPCell bar1 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                            bar1.HorizontalAlignment = Element.ALIGN_LEFT;
                            bar1.Border = 0;
                            bar1.Colspan = 8;
                            bar1.BackgroundColor = new BaseColor(135, 0, 27);
                            HeadT.AddCell(bar1);

                            PdfPCell stinfo = new PdfPCell(new Phrase("Student Name: " + stcomm[1], new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD, BaseColor.BLACK)));
                            stinfo.HorizontalAlignment = Element.ALIGN_LEFT;
                            stinfo.Border = 0;
                            stinfo.Colspan = 4;
                            stinfo.PaddingTop = 3;
                            HeadT.AddCell(stinfo);

                            PdfPCell messag = new PdfPCell(new Phrase("This report describes progress toward grade level learning expectations, identifies successes and provides guidance for improvement.", new Font(Font.FontFamily.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                            messag.HorizontalAlignment = Element.ALIGN_LEFT;
                            messag.Border = 0;
                            messag.Rowspan = 3;
                            messag.Colspan = 3;
                            messag.PaddingTop = 5;
                            messag.PaddingBottom = 5;
                            HeadT.AddCell(messag);

                            PdfPCell legd = new PdfPCell(new Phrase("A = Absence" + Environment.NewLine + "T = Tardies", new Font(Font.FontFamily.HELVETICA, 9, Font.NORMAL, BaseColor.BLACK)));
                            legd.HorizontalAlignment = Element.ALIGN_LEFT;
                            legd.VerticalAlignment = Element.ALIGN_BOTTOM;
                            legd.Border = 0;
                            legd.Rowspan = 3;
                            legd.PaddingTop = 5;
                            legd.PaddingLeft = 20;
                            legd.PaddingBottom = 5;
                            HeadT.AddCell(legd);

                            PdfPCell grade = new PdfPCell(new Phrase("Grade: " + stcomm[2], new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD, BaseColor.BLACK)));
                            grade.HorizontalAlignment = Element.ALIGN_LEFT;
                            grade.Border = 0;
                            grade.Colspan = 4;
                            HeadT.AddCell(grade);

                            PdfPCell boh = new PdfPCell(new Phrase("Bohio: " + stcomm[4], new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD, BaseColor.BLACK)));
                            boh.HorizontalAlignment = Element.ALIGN_LEFT;
                            boh.Border = 0;
                            boh.Colspan = 3;
                            boh.PaddingBottom = 5;
                            HeadT.AddCell(boh);


                            PdfPCell stid = new PdfPCell(new Phrase("StudentID: " + stcomm[0], new Font(Font.FontFamily.HELVETICA, 8, Font.BOLD, BaseColor.WHITE)));
                            stid.HorizontalAlignment = Element.ALIGN_LEFT;
                            stid.Border = 0;
                            stid.PaddingBottom = 5;
                            HeadT.AddCell(stid);


                            PdfPTable GradeTable = new PdfPTable(18);
                            GradeTable.HorizontalAlignment = Element.ALIGN_CENTER;
                            GradeTable.WidthPercentage = 100;

                            PdfPCell Course = new PdfPCell(new Phrase("Course / Teacher", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                            Course.HorizontalAlignment = Element.ALIGN_LEFT;
                            Course.BackgroundColor = new BaseColor(135, 0, 27);
                            Course.BorderWidth = 1F;
                            Course.Colspan = 5;
                            Course.PaddingBottom = 3;
                            Course.PaddingLeft = 3;
                            GradeTable.AddCell(Course);
                           
                            PdfPCell Q1 = new PdfPCell(new Phrase("Q1", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                            Q1.HorizontalAlignment = Element.ALIGN_CENTER;
                            Q1.BackgroundColor = new BaseColor(135, 0, 27);
                            Q1.BorderWidth = 1F;
                            Q1.PaddingBottom = 3;
                            Q1.PaddingLeft = 3;
                            GradeTable.AddCell(Q1);

                            PdfPCell CommS1 = new PdfPCell(new Phrase("Comment", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                            CommS1.HorizontalAlignment = Element.ALIGN_LEFT;
                            CommS1.BackgroundColor = new BaseColor(135, 0, 27);
                            CommS1.BorderWidth = 1F;
                            CommS1.Colspan = 11;
                            CommS1.PaddingBottom = 3;
                            CommS1.PaddingLeft = 3;
                            GradeTable.AddCell(CommS1);
                            PdfPCell ABS = new PdfPCell(new Phrase("A/T", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                            ABS.HorizontalAlignment = Element.ALIGN_CENTER;
                            ABS.BackgroundColor = new BaseColor(135, 0, 27);
                            ABS.BorderWidth = 1F;
                            ABS.PaddingBottom = 3;
                            ABS.PaddingLeft = 3;
                            GradeTable.AddCell(ABS);




                            PdfPCell CO;
                            PdfPCell TE;
                            PdfPCell QU1;
                            PdfPCell COM1;
                            PdfPCell ABS1;


                            for (int i = 0; i < stTable.Length - 1; i++)
                            {
                                var nfila = stTable[i].Split('|');
                                //+ Environment.NewLine + nfila[1]
                                CO = new PdfPCell(new Phrase(nfila[0], new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.BLACK)));
                                CO.HorizontalAlignment = Element.ALIGN_LEFT;
                                CO.BorderWidth = 0.5F;
                                CO.Colspan = 5;
                                CO.PaddingTop = 5;
                                CO.PaddingLeft = 5;
                                CO.BorderWidthBottom = 0;
                                CO.BorderColor = BaseColor.GRAY;
                                GradeTable.AddCell(CO);

                                QU1 = new PdfPCell(new Phrase(nfila[2], new Font(Font.FontFamily.HELVETICA, 11.0F, Font.NORMAL, BaseColor.BLACK)));
                                QU1.HorizontalAlignment = Element.ALIGN_CENTER;
                                QU1.BorderWidth = 0.5F;
                                QU1.BorderColor = BaseColor.GRAY;
                                QU1.Rowspan = 2;
                                QU1.PaddingTop = 5;

                                GradeTable.AddCell(QU1);
                                COM1 = new PdfPCell(new Phrase(nfila[3], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                                COM1.HorizontalAlignment = Element.ALIGN_LEFT;
                                COM1.BorderWidth = 0.5F;
                                COM1.BorderColor = BaseColor.GRAY;
                                COM1.Colspan = 11;
                                COM1.Rowspan = 2;
                                COM1.PaddingBottom = 5;
                                COM1.PaddingLeft = 5;
                                GradeTable.AddCell(COM1);
                                ABS1 = new PdfPCell(new Phrase(nfila[4] + '/' + nfila[5], new Font(Font.FontFamily.HELVETICA, 11.0F, Font.NORMAL, BaseColor.BLACK)));
                                ABS1.HorizontalAlignment = Element.ALIGN_CENTER;
                                ABS1.BorderWidth = 0.5F;
                                ABS1.Rowspan = 2;
                                ABS1.PaddingTop = 5;
                                ABS1.BorderColor = BaseColor.GRAY;
                                GradeTable.AddCell(ABS1);
                                TE = new PdfPCell(new Phrase(nfila[1], new Font(Font.FontFamily.HELVETICA, 11.0F, Font.ITALIC, BaseColor.BLACK)));
                                TE.HorizontalAlignment = Element.ALIGN_LEFT;
                                TE.Colspan = 5;
                                TE.BorderWidth = 0.5F;
                                TE.PaddingLeft = 5;
                                TE.BorderWidthTop = 0;
                                TE.BorderColor = BaseColor.GRAY;
                                GradeTable.AddCell(TE);


                            }

                            PdfPTable FOOT = new PdfPTable(8);
                            FOOT.HorizontalAlignment = Element.ALIGN_LEFT;
                            FOOT.WidthPercentage = 100;

                           
                            PdfPCell Comms = new PdfPCell(new Phrase("Community Service Hours", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.BLACK)));
                            Comms.HorizontalAlignment = Element.ALIGN_LEFT;
                            Comms.Colspan = 8;
                            Comms.Border = 0;
                            Comms.PaddingTop = 5;
                            FOOT.AddCell(Comms);

                            PdfPCell Curr = new PdfPCell(new Phrase("Current Year Hours: " + stcomm[5], new Font(Font.FontFamily.HELVETICA, 9.0F, Font.NORMAL, BaseColor.BLACK)));
                            Curr.HorizontalAlignment = Element.ALIGN_LEFT;
                            Curr.Colspan = 3;
                            Curr.Border = 0;
                            FOOT.AddCell(Curr);

                            PdfPCell hr = new PdfPCell(new Phrase("- 15 hours required each year (minimum of 10 outreach)", new Font(Font.FontFamily.HELVETICA, 9.0F, Font.NORMAL, BaseColor.BLACK)));
                            hr.HorizontalAlignment = Element.ALIGN_LEFT;
                            hr.Colspan = 5;
                            hr.Border = 0;
                            FOOT.AddCell(hr);

                            PdfPCell InH = new PdfPCell(new Phrase("    In School: " + stcomm[6], new Font(Font.FontFamily.HELVETICA, 9.0F, Font.NORMAL, BaseColor.BLACK)));
                            InH.HorizontalAlignment = Element.ALIGN_LEFT;
                            InH.Colspan = 3;
                            InH.Border = 0;
                            FOOT.AddCell(InH);

                            PdfPCell hr2 = new PdfPCell(new Phrase("- 60 hours minimum required in grades 9 to 12", new Font(Font.FontFamily.HELVETICA, 9.0F, Font.NORMAL, BaseColor.BLACK)));
                            hr2.HorizontalAlignment = Element.ALIGN_LEFT;
                            hr2.Colspan = 5;
                            hr2.Border = 0;
                            FOOT.AddCell(hr2);

                            PdfPCell outH = new PdfPCell(new Phrase("    Outreach: " + stcomm[7], new Font(Font.FontFamily.HELVETICA, 9.0F, Font.NORMAL, BaseColor.BLACK)));
                            outH.HorizontalAlignment = Element.ALIGN_LEFT;
                            outH.Colspan = 8;
                            outH.Border = 0;
                            FOOT.AddCell(outH);

                            PdfPCell Total = new PdfPCell(new Phrase("Total (Grade 9-12) Hours: " + stcomm[8], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.BLACK)));
                            Total.HorizontalAlignment = Element.ALIGN_LEFT;
                            Total.Colspan = 8;
                            Total.Border = 0;
                            FOOT.AddCell(Total);
                                             
                            PdfPCell IFY = new PdfPCell(new Phrase("If you have questions regarding this progress report, please contact the High School Office: 809-947-1033.", new Font(Font.FontFamily.HELVETICA, 8.5F, Font.NORMAL, BaseColor.BLACK)));
                            IFY.HorizontalAlignment = Element.ALIGN_LEFT;
                            IFY.Colspan = 8;
                            IFY.Border = 0;
                            IFY.PaddingTop = 5;
                            FOOT.AddCell(IFY);

                    
                            
                            //documento.Add(stfoto);
                            documento.Add(HeadT);
                            documento.Add(GradeTable);
                            documento.Add(FOOT);

                            //Process prc = new System.Diagnostics.Process();
                            //prc.StartInfo.FileName = fileName;
                            //prc.Start();
                        }
                        else
                        {
                            con.Close();

                        }
                        documento.NewPage();
                    }

                    con.Close();

                }
                else
                {
                    if (con.State == ConnectionState.Closed)
                    {
                        con.Open();
                    }
                    string stdata = string.Empty;
                    string datat = string.Empty;
                    fname = "HS_ProgressReport_" + DateTime.Now.DayOfYear + DateTime.Now.Month + DateTime.Now.Day + DateTime.Now.Millisecond + ".pdf";
                    fileName = HttpContext.Current.Server.MapPath("~/RepoFiles/" + fname);
                    PdfWriter.GetInstance(documento, new FileStream(fileName, FileMode.Create));
                    documento.Open();

                    sql = "WITH main_query AS ( SELECT C.COURSE_NAME,T.FIRST_NAME||' '||T.LAST_NAME AS TEACHER, PG.FINALGRADENAME,PG.GRADE,";
                    sql += " TO_CHAR(PG.COMMENT_VALUE) AS COMMENTS,CO.CURRENTABSENCES,CO.CURRENTTARDIES FROM CC CO";
                    sql += " LEFT JOIN STUDENTS S ON CO.STUDENTID=S.ID";
                    sql += " LEFT JOIN COURSES C ON CO.COURSE_NUMBER=C.COURSE_NUMBER";
                    sql += " LEFT JOIN TEACHERS T ON CO.TEACHERID=T.ID";
                    sql += " LEFT JOIN PGFINALGRADES PG ON CO.SECTIONID=PG.SECTIONID AND CO.STUDENTID=PG.STUDENTID";
                    sql += " WHERE CO.TERMID IN(2800,2801,2802) AND C.COURSE_NAME NOT LIKE '%Bohio%' AND PG.FINALGRADENAME IN ('Q1','S1') AND S.STUDENT_NUMBER=" + stnum + "";
                    sql += " )";
                    sql += " SELECT DISTINCT COURSE_NAME,TEACHER";
                    sql += " ,(SELECT  y.GRADE FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.FINALGRADENAME='Q1') Q1";
                    sql += " ,(SELECT  y.COMMENTS FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.FINALGRADENAME='Q1') Comments";
                    sql += " ,(SELECT  y.CURRENTABSENCES FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.FINALGRADENAME='Q1') ABS1";
                    sql += " ,(SELECT  y.CURRENTTARDIES FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.FINALGRADENAME='Q1') TARDI";
                    sql += "  FROM	main_query M";
                    sql += " ORDER BY COURSE_NAME";

                    OracleCommand cmd = new OracleCommand(sql, con);
                    OracleDataReader odr = cmd.ExecuteReader();
                    while (odr.Read())
                    {
                        datat += odr["COURSE_NAME"].ToString() + '|';
                        datat += odr["TEACHER"].ToString() + '|';
                        datat += odr["Q1"].ToString() + '|';
                        datat += odr["Comments"].ToString() + '|';
                        datat += odr["ABS1"].ToString() + '|';
                        datat += odr["TARDI"].ToString() + '^';

                    }
                    if (datat != "")
                    {
                        // Close and Dispose OracleConnection object
                        sql = "SELECT S.STUDENT_NUMBER,S.FIRST_NAME||' '||S.LAST_NAME AS STUDENT,S.GRADE_LEVEL,C.COURSE_NAME ,T.FIRST_NAME||' '||T.LAST_NAME AS TEACHER,CO.HS_SERVICE_HOURS_CURRENT,CO.HS_SERVICE_HOURS_CURRENT_IN,";
                        sql += " CO.HS_SERVICE_HOURS_CURRENT_OUTRE,CO.HS_TOTAL_SERVICE_HOURS  FROM CC CO";
                        sql += " LEFT JOIN STUDENTS S ON CO.STUDENTID=S.ID";
                        sql += " LEFT JOIN COURSES C ON CO.COURSE_NUMBER=C.COURSE_NUMBER";
                        sql += " LEFT JOIN TEACHERS T ON CO.TEACHERID=T.ID";
                        sql += " LEFT JOIN U_COMMUNITYSERVICE CO ON S.DCID=CO.STUDENTSDCID";
                        sql += " WHERE CO.TERMID IN(2800) AND COURSE_NAME LIKE '%Bohio%' AND S.STUDENT_NUMBER=" + stnum + "";

                        OracleCommand cmd1 = new OracleCommand(sql, con);
                        OracleDataReader odr1 = cmd1.ExecuteReader();
                        while (odr1.Read())
                        {
                            stdata += odr1["STUDENT_NUMBER"].ToString() + '|';
                            stdata += odr1["STUDENT"].ToString() + '|';
                            stdata += odr1["GRADE_LEVEL"].ToString() + '|';
                            stdata += odr1["COURSE_NAME"].ToString() + '|';
                            stdata += odr1["TEACHER"].ToString() + '|';
                            stdata += odr1["HS_SERVICE_HOURS_CURRENT"].ToString() + '|';
                            stdata += odr1["HS_SERVICE_HOURS_CURRENT_IN"].ToString() + '|';
                            stdata += odr1["HS_SERVICE_HOURS_CURRENT_OUTRE"].ToString() + '|';
                            stdata += odr1["HS_TOTAL_SERVICE_HOURS"].ToString() + '|';

                        }

                        var stTable = datat.Split('^');
                        var stcomm = stdata.Split('|');

                        iTextSharp.text.Image Imagen = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("~/img/WLOGO.jpg"));
                        // Imagen.SetAbsolutePosition(-3, 520);
                        Imagen.ScalePercent(2.5f);



                        PdfPTable HeadT = new PdfPTable(8);
                        HeadT.HorizontalAlignment = Element.ALIGN_CENTER;
                        HeadT.WidthPercentage = 100;

                        PdfPCell logo = new PdfPCell(Imagen);
                        logo.Colspan = 4;
                        logo.Border = 0;
                        logo.HorizontalAlignment = Element.ALIGN_LEFT;
                        logo.Rowspan = 3;
                        logo.Padding = 3;
                        HeadT.AddCell(logo);

                                     

                        PdfPCell HS = new PdfPCell(new Phrase("HIGH SCHOOL Progress Report", new Font(Font.FontFamily.HELVETICA, 14, Font.BOLD, new BaseColor(135, 0, 27))));
                        HS.HorizontalAlignment = Element.ALIGN_LEFT;
                        HS.Colspan = 4;
                        HS.Border = 0;
                        HeadT.AddCell(HS);
                      
                        PdfPCell SQ1 = new PdfPCell(new Phrase("School Year 2018-19 Midsemester 1 (Q1)", new Font(Font.FontFamily.HELVETICA, 14, Font.BOLD, BaseColor.BLACK)));
                        SQ1.HorizontalAlignment = Element.ALIGN_LEFT;
                        SQ1.Colspan = 4;
                        SQ1.Border = 0;
                        HeadT.AddCell(SQ1);

                        PdfPCell Pub = new PdfPCell(new Phrase("Published " + DateTime.Now.ToString("MMMM dd, yyyy"), new Font(Font.FontFamily.HELVETICA, 12, Font.ITALIC, BaseColor.BLACK)));
                        Pub.Colspan = 4;
                        Pub.HorizontalAlignment = Element.ALIGN_LEFT;
                        Pub.Border = 0;
                        HeadT.AddCell(Pub);

                        PdfPCell bar1 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                        bar1.HorizontalAlignment = Element.ALIGN_LEFT;
                        bar1.Border = 0;
                        bar1.Colspan = 8;
                        bar1.BackgroundColor = new BaseColor(135, 0, 27);
                        HeadT.AddCell(bar1);

                        PdfPCell stinfo = new PdfPCell(new Phrase("Student Name: " + stcomm[1], new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD, BaseColor.BLACK)));
                        stinfo.HorizontalAlignment = Element.ALIGN_LEFT;
                        stinfo.Border = 0;
                        stinfo.Colspan = 4;
                        stinfo.PaddingTop = 3;
                        HeadT.AddCell(stinfo);

                        PdfPCell messag = new PdfPCell(new Phrase("This report describes progress toward grade level learning expectations, identifies successes and provides guidance for improvement.", new Font(Font.FontFamily.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                        messag.HorizontalAlignment = Element.ALIGN_LEFT;
                        messag.Border = 0;
                        messag.Rowspan = 3;
                        messag.Colspan = 3;
                        messag.PaddingTop = 5;
                        messag.PaddingBottom = 5;
                        HeadT.AddCell(messag);
                        PdfPCell legd = new PdfPCell(new Phrase("A = Absence"+Environment.NewLine+"T = Tardies", new Font(Font.FontFamily.HELVETICA, 9, Font.NORMAL, BaseColor.BLACK)));
                        legd.HorizontalAlignment = Element.ALIGN_LEFT;
                        legd.VerticalAlignment = Element.ALIGN_BOTTOM;
                        legd.Border = 0;
                        legd.Rowspan = 3;
                        legd.PaddingTop = 5;
                        legd.PaddingLeft = 20;
                        legd.PaddingBottom = 5;
                        HeadT.AddCell(legd);

                        PdfPCell grade = new PdfPCell(new Phrase("Grade: " + stcomm[2], new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD, BaseColor.BLACK)));
                        grade.HorizontalAlignment = Element.ALIGN_LEFT;
                        grade.Border = 0;
                        grade.Colspan = 4;
                        HeadT.AddCell(grade);

                        PdfPCell boh = new PdfPCell(new Phrase("Bohio: " + stcomm[4], new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD, BaseColor.BLACK)));
                        boh.HorizontalAlignment = Element.ALIGN_LEFT;
                        boh.Border = 0;
                        boh.Colspan = 3;
                        boh.PaddingBottom = 5;
                        HeadT.AddCell(boh);


                        PdfPCell stid = new PdfPCell(new Phrase("StudentID: " + stcomm[0], new Font(Font.FontFamily.HELVETICA, 8, Font.BOLD, BaseColor.WHITE)));
                        stid.HorizontalAlignment = Element.ALIGN_LEFT;
                        stid.Border = 0;
                        stid.PaddingBottom = 5;
                        
                        HeadT.AddCell(stid);
                                              
                    
                        PdfPTable GradeTable = new PdfPTable(18);
                        GradeTable.HorizontalAlignment = Element.ALIGN_CENTER;
                        GradeTable.WidthPercentage = 100;

                        PdfPCell Course = new PdfPCell(new Phrase("Course / Teacher", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                        Course.HorizontalAlignment = Element.ALIGN_LEFT;
                        Course.BackgroundColor = new BaseColor(135, 0, 27);
                        Course.BorderWidth = 1F;
                        Course.Colspan = 5;
                        Course.PaddingBottom = 3;
                        Course.PaddingLeft = 3;
                        GradeTable.AddCell(Course);
                     
                        PdfPCell Q1 = new PdfPCell(new Phrase("Q1", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                        Q1.HorizontalAlignment = Element.ALIGN_CENTER;
                        Q1.BackgroundColor = new BaseColor(135, 0, 27);
                        Q1.BorderWidth = 1F;
                        Q1.PaddingBottom = 3;
                        Q1.PaddingLeft = 3;
                        GradeTable.AddCell(Q1);

                        PdfPCell CommS1 = new PdfPCell(new Phrase("Comment", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                        CommS1.HorizontalAlignment = Element.ALIGN_LEFT;
                        CommS1.BackgroundColor = new BaseColor(135, 0, 27);
                        CommS1.BorderWidth = 1F;
                        CommS1.Colspan = 11;
                        CommS1.PaddingBottom = 3;
                        CommS1.PaddingLeft = 3;
                        GradeTable.AddCell(CommS1);
                        PdfPCell ABS = new PdfPCell(new Phrase("A/T", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.WHITE)));
                        ABS.HorizontalAlignment = Element.ALIGN_CENTER;
                        ABS.BackgroundColor = new BaseColor(135, 0, 27);
                        ABS.BorderWidth = 1F;
                        ABS.PaddingBottom = 3;
                        ABS.PaddingLeft = 3;
                        GradeTable.AddCell(ABS);
                       



                        PdfPCell CO;
                        PdfPCell TE;
                        PdfPCell QU1;
                        PdfPCell COM1;
                        PdfPCell ABS1;
                        

                        for (int i = 0; i < stTable.Length - 1; i++)
                        {
                            var nfila = stTable[i].Split('|');
                            //+ Environment.NewLine + nfila[1]
                            CO = new PdfPCell(new Phrase(nfila[0], new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.BLACK)));
                            CO.HorizontalAlignment = Element.ALIGN_LEFT;
                            CO.BorderWidth = 0.5F;
                            CO.Colspan = 5;
                            CO.PaddingTop = 5;
                            CO.PaddingLeft = 5;
                            CO.BorderWidthBottom = 0;
                            CO.BorderColor = BaseColor.GRAY;
                            GradeTable.AddCell(CO);
                           
                            QU1 = new PdfPCell(new Phrase(nfila[2], new Font(Font.FontFamily.HELVETICA, 11.0F, Font.NORMAL, BaseColor.BLACK)));
                            QU1.HorizontalAlignment = Element.ALIGN_CENTER;
                            QU1.BorderWidth = 0.5F;
                            QU1.BorderColor = BaseColor.GRAY;
                            QU1.Rowspan = 2;
                            QU1.PaddingTop = 5;
                            
                            GradeTable.AddCell(QU1);
                            COM1 = new PdfPCell(new Phrase(nfila[3], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                            COM1.HorizontalAlignment = Element.ALIGN_LEFT;
                            COM1.BorderWidth = 0.5F;
                            COM1.BorderColor = BaseColor.GRAY;
                            COM1.Colspan = 11;
                            COM1.Rowspan = 2;
                            COM1.PaddingBottom = 5;
                            COM1.PaddingLeft = 5;
                            GradeTable.AddCell(COM1);
                            ABS1 = new PdfPCell(new Phrase(nfila[4]+'/'+ nfila[5], new Font(Font.FontFamily.HELVETICA, 11.0F, Font.NORMAL, BaseColor.BLACK)));
                            ABS1.HorizontalAlignment = Element.ALIGN_CENTER;
                            ABS1.BorderWidth = 0.5F;
                            ABS1.Rowspan = 2;
                            ABS1.PaddingTop = 5;
                            ABS1.BorderColor = BaseColor.GRAY;
                            GradeTable.AddCell(ABS1);
                            TE = new PdfPCell(new Phrase(nfila[1], new Font(Font.FontFamily.HELVETICA, 11.0F, Font.ITALIC, BaseColor.BLACK)));
                            TE.HorizontalAlignment = Element.ALIGN_LEFT;
                            TE.Colspan = 5;
                            TE.BorderWidth = 0.5F;
                            TE.PaddingLeft = 5;
                            TE.BorderWidthTop = 0;
                            TE.BorderColor = BaseColor.GRAY;
                            GradeTable.AddCell(TE);


                        }


                        PdfPTable FOOT = new PdfPTable(8);
                        FOOT.HorizontalAlignment = Element.ALIGN_LEFT;
                        FOOT.WidthPercentage = 100;

                                                
                        PdfPCell Comms = new PdfPCell(new Phrase("Community Service Hours", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.BLACK)));
                        Comms.HorizontalAlignment = Element.ALIGN_LEFT;
                        Comms.Colspan = 8;
                        Comms.Border = 0;
                        Comms.PaddingTop = 5;
                        FOOT.AddCell(Comms);

                        PdfPCell Curr = new PdfPCell(new Phrase("Current Year Hours: " + stcomm[5], new Font(Font.FontFamily.HELVETICA, 9.0F, Font.NORMAL, BaseColor.BLACK)));
                        Curr.HorizontalAlignment = Element.ALIGN_LEFT;
                        Curr.Colspan = 3;
                        Curr.Border = 0;
                        FOOT.AddCell(Curr);

                        PdfPCell hr = new PdfPCell(new Phrase("- 15 hours required each year (minimum of 10 outreach)", new Font(Font.FontFamily.HELVETICA, 9.0F, Font.NORMAL, BaseColor.BLACK)));
                        hr.HorizontalAlignment = Element.ALIGN_LEFT;
                        hr.Colspan = 5;
                        hr.Border = 0;
                        FOOT.AddCell(hr);

                        PdfPCell InH = new PdfPCell(new Phrase("    In School: " + stcomm[6], new Font(Font.FontFamily.HELVETICA, 9.0F, Font.NORMAL, BaseColor.BLACK)));
                        InH.HorizontalAlignment = Element.ALIGN_LEFT;
                        InH.Colspan = 3;
                        InH.Border = 0;
                        FOOT.AddCell(InH);

                        PdfPCell hr2 = new PdfPCell(new Phrase("- 60 hours minimum required in grades 9 to 12", new Font(Font.FontFamily.HELVETICA, 9.0F, Font.NORMAL, BaseColor.BLACK)));
                        hr2.HorizontalAlignment = Element.ALIGN_LEFT;
                        hr2.Colspan = 5;
                        hr2.Border = 0;
                        FOOT.AddCell(hr2);

                        PdfPCell outH = new PdfPCell(new Phrase("    Outreach: " + stcomm[7], new Font(Font.FontFamily.HELVETICA, 9.0F, Font.NORMAL, BaseColor.BLACK)));
                        outH.HorizontalAlignment = Element.ALIGN_LEFT;
                        outH.Colspan = 8;
                        outH.Border = 0;
                        FOOT.AddCell(outH);

                        PdfPCell Total = new PdfPCell(new Phrase("Total (Grade 9-12) Hours: " + stcomm[8], new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.BLACK)));
                        Total.HorizontalAlignment = Element.ALIGN_LEFT;
                        Total.Colspan = 8;
                        Total.Border = 0;
                        FOOT.AddCell(Total);

                        PdfPCell IFY = new PdfPCell(new Phrase("If you have questions regarding this progress report, please contact the High School Office: 809-947-1033.", new Font(Font.FontFamily.HELVETICA, 8.5F, Font.NORMAL, BaseColor.BLACK)));
                        IFY.HorizontalAlignment = Element.ALIGN_LEFT;
                        IFY.Colspan = 8;
                        IFY.Border = 0;
                        IFY.PaddingTop = 5;
                        FOOT.AddCell(IFY);



                        //documento.Add(stfoto);
                        documento.Add(HeadT);
                        documento.Add(GradeTable);
                        documento.Add(FOOT);



                        //Process prc = new System.Diagnostics.Process();
                        //prc.StartInfo.FileName = fileName;
                        //prc.Start();
                    }
                    else
                    {
                        con.Close();
                        fname = "";
                    }
                    con.Close();

                }

                documento.Close();
                con.Dispose();
            }
            catch (Exception ex)
            {
                throw;
            }
            return fname;
        }

        
        [WebMethod]
        public static string reportCard(string stnum)
        {
            string sql = string.Empty;
           // int stnum = 4388;
           
           
            string fname = string.Empty;
            string fileName = string.Empty;
            OracleConnection con = new OracleConnection();
            con.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings["conn"].ConnectionString;
            Document documento = new Document(PageSize.LETTER.Rotate(), 10, 10, 5, 5);
            try
            {
                


                if (stnum.IndexOf(';') > -1)
                {
                    var stnumb = stnum.Split(';');
                    fname = "HS_ReportCardS1_" + DateTime.Now.DayOfYear + DateTime.Now.Month + DateTime.Now.Day + DateTime.Now.Millisecond + ".pdf";
                    fileName = HttpContext.Current.Server.MapPath("~/RepoFiles/" + fname);
                    PdfWriter.GetInstance(documento, new FileStream(fileName, FileMode.Create));
                    documento.Open();
                   
                    for (int a = 0; a < stnumb.Length; a++)
                    {
                        string stdata = string.Empty;
                        string S1GPA = string.Empty;
                        string CUGPA = string.Empty;
                        string datat = string.Empty;
                        if (con.State == ConnectionState.Closed)
                        {
                            con.Open();
                        }
                        
                        sql = "WITH main_query AS(SELECT DISTINCT S.STUDENT_NUMBER, SG.COURSE_NAME, SG.STORECODE, SG.GRADE, SG.TEACHER_NAME,";
                    sql += " TO_CHAR(SG.COMMENT_VALUE) AS COMMENTS, SG.ABSENCES, SG.TARDIES FROM   STOREDGRADES SG";
                    sql += " LEFT JOIN STUDENTS S ON SG.STUDENTID = S.ID";
                    sql += " WHERE SG.TERMID IN(2700, 2701)  AND STORECODE IN('Q1', 'Q2', 'E1', 'S1') AND S.STUDENT_NUMBER =" + stnumb[a] + "";
                    sql += " )";
                    sql += " SELECT DISTINCT COURSE_NAME, TEACHER_NAME";
                    sql += " ,(SELECT  y.GRADE FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.STORECODE = 'Q1') Q1";
                    sql += " ,(SELECT  y.GRADE FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.STORECODE = 'Q2') Q2";
                    sql += " ,(SELECT  y.GRADE FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.STORECODE = 'E1') E1";
                    sql += " ,(SELECT  y.GRADE FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.STORECODE = 'S1') S1";
                    sql += " ,(SELECT  y.COMMENTS FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.STORECODE = 'S1') Comments";
                    sql += " ,(SELECT  y.ABSENCES FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.STORECODE = 'S1') ABS1";
                    sql += " ,(SELECT  y.TARDIES FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.STORECODE = 'S1') TARDI";
                    sql += "  FROM main_query M";
                    sql += " ORDER BY COURSE_NAME";

                    OracleCommand cmd = new OracleCommand(sql, con);
                    OracleDataReader odr = cmd.ExecuteReader();
                    while (odr.Read())
                    {
                        datat += odr["COURSE_NAME"].ToString() + '|';
                        datat += odr["TEACHER_NAME"].ToString() + '|';
                        datat += odr["Q1"].ToString() + '|';
                        datat += odr["Q2"].ToString() + '|';
                        datat += odr["E1"].ToString() + '|';
                        datat += odr["S1"].ToString() + '|';
                        datat += odr["Comments"].ToString() + '|';
                        datat += odr["ABS1"].ToString() + '|';
                        datat += odr["TARDI"].ToString() + '^';

                    }
                    if (datat != "")
                    {
                        // Close and Dispose OracleConnection object

                        sql = "SELECT DISTINCT T.ABBREVIATION,s.student_number,s.lastfirst,s.grade_level,S.studentpict_guid,CO.HS_SERVICE_HOURS_CURRENT,CO.HS_SERVICE_HOURS_CURRENT_IN,CO.HS_SERVICE_HOURS_CURRENT_OUTRE,CO.HS_TOTAL_SERVICE_HOURS FROM STUDENTS S";
                        sql += " LEFT JOIN U_COMMUNITYSERVICE CO ON S.DCID=CO.STUDENTSDCID";
                        sql += " LEFT JOIN STOREDGRADES SG ON S.ID=SG.STUDENTID";
                        sql += " LEFT JOIN TERMS T ON SG.TERMID=T.ID";
                        sql += " WHERE SG.TERMID IN (2700) AND S.STUDENT_NUMBER =" + stnumb[a] + "";

                        OracleCommand cmd1 = new OracleCommand(sql, con);
                        OracleDataReader odr1 = cmd1.ExecuteReader();
                        while (odr1.Read())
                        {
                            stdata += odr1["ABBREVIATION"].ToString() + '|';
                            stdata += odr1["student_number"].ToString() + '|';
                            stdata += odr1["lastfirst"].ToString() + '|';
                            stdata += odr1["grade_level"].ToString() + '|';
                            stdata += odr1["studentpict_guid"].ToString() + '|';
                            stdata += odr1["HS_SERVICE_HOURS_CURRENT"].ToString() + '|';
                            stdata += odr1["HS_SERVICE_HOURS_CURRENT_IN"].ToString() + '|';
                            stdata += odr1["HS_SERVICE_HOURS_CURRENT_OUTRE"].ToString() + '|';
                            stdata += odr1["HS_TOTAL_SERVICE_HOURS"].ToString() + '|';

                        }

                        sql = "SELECT S.STUDENT_NUMBER, ROUND(SUM(sg.gpa_points)/COUNT(sg.gpa_points),3) AS GPA FROM STOREDGRADES SG";
                        sql += " LEFT JOIN STUDENTS S ON SG.STUDENTID=S.ID";
                        sql += " WHERE SG.TERMID IN (2700,2702,2701) AND S.STUDENT_NUMBER =" + stnumb[a] + "";
                        sql += " AND SG.STORECODE IN ('S1') AND SG.GPA_POINTS<>0 ";
                        sql += " GROUP BY S.STUDENT_NUMBER";


                        OracleCommand cmd2 = new OracleCommand(sql, con);
                        OracleDataReader odr2 = cmd2.ExecuteReader();
                        while (odr2.Read())
                        {
                            S1GPA += odr2["STUDENT_NUMBER"].ToString() + '|';
                            S1GPA += odr2["GPA"].ToString() + '|';

                        }

                        sql = "SELECT S.STUDENT_NUMBER, ROUND(SUM(sg.gpa_points)/COUNT(sg.gpa_points),3) AS GPA FROM STOREDGRADES SG";
                        sql += " LEFT JOIN STUDENTS S ON SG.STUDENTID=S.ID";
                        sql += " WHERE SG.TERMID IN (2700,2702,2701,2600,2602,2601,2500,2502,2501) AND S.STUDENT_NUMBER =" + stnumb[a] + "";
                        sql += " AND SG.STORECODE IN ('S1','S2') AND SG.GPA_POINTS<>0 ";
                        sql += " GROUP BY S.STUDENT_NUMBER";


                        OracleCommand cmd3 = new OracleCommand(sql, con);
                        OracleDataReader odr3 = cmd3.ExecuteReader();
                        while (odr3.Read())
                        {
                            CUGPA += odr3["STUDENT_NUMBER"].ToString() + '|';
                            CUGPA += odr3["GPA"].ToString() + '|';

                        }

                        // Close and Dispose OracleConnection object
                       

                        var stTable = datat.Split('^');
                        var stcomm = stdata.Split('|');
                        var stS1GPA = S1GPA.Split('|');
                        var stCUGPA = CUGPA.Split('|');

                       
                            
                      
                            
                        iTextSharp.text.Image Imagen = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("~/img/cmslogo.jpg"));
                        Imagen.SetAbsolutePosition(-3, 520);
                        Imagen.ScalePercent(2f);

                        //byte[] imageBytes = Convert.FromBase64String(@"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LyIgeGFwUmlnaHRzOk1hcmtlZD0iVHJ1ZSIgeGFwUmlnaHRzOldlYlN0YXRlbWVudD0iaHR0cDovL3Byby5jb3JiaXMuY29tL3NlYXJjaC9zZWFyY2hyZXN1bHRzLmFzcD90eHQ9NDItMTU1NjQ5NzgmYW1wO29wZW5JbWFnZT00Mi0xNTU2NDk3OCI+DQoJCQk8ZGM6cmlnaHRzPg0KCQkJCTxyZGY6QWx0Pg0KCQkJCQk8cmRmOmxpIHhtbDpsYW5nPSJ4LWRlZmF1bHQiPsKpIENvcmJpcy4gIEFsbCBSaWdodHMgUmVzZXJ2ZWQuPC9yZGY6bGk+DQoJCQkJPC9yZGY6QWx0Pg0KCQkJPC9kYzpyaWdodHM+DQoJCQk8ZGM6Y3JlYXRvcj48cmRmOlNlcSB4bWxuczpyZGY9Imh0dHA6Ly93d3cudzMub3JnLzE5OTkvMDIvMjItcmRmLXN5bnRheC1ucyMiPjxyZGY6bGk+Q29yYmlzPC9yZGY6bGk+PC9yZGY6U2VxPg0KCQkJPC9kYzpjcmVhdG9yPjwvcmRmOkRlc2NyaXB0aW9uPg0KCQk8cmRmOkRlc2NyaXB0aW9uIHhtbG5zOnhtcD0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wLyI+PHhtcDpSYXRpbmc+NDwveG1wOlJhdGluZz48eG1wOkNyZWF0ZURhdGU+MjAwOC0wMi0xMVQxOTozMjo0My4xNzNaPC94bXA6Q3JlYXRlRGF0ZT48L3JkZjpEZXNjcmlwdGlvbj48cmRmOkRlc2NyaXB0aW9uIHhtbG5zOk1pY3Jvc29mdFBob3RvPSJodHRwOi8vbnMubWljcm9zb2Z0LmNvbS9waG90by8xLjAvIj48TWljcm9zb2Z0UGhvdG86UmF0aW5nPjYzPC9NaWNyb3NvZnRQaG90bzpSYXRpbmc+PC9yZGY6RGVzY3JpcHRpb24+PC9yZGY6UkRGPg0KPC94OnhtcG1ldGE+DQogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICA8P3hwYWNrZXQgZW5kPSd3Jz8+/9sAQwACAQECAQECAgICAgICAgMFAwMDAwMGBAQDBQcGBwcHBgcHCAkLCQgICggHBwoNCgoLDAwMDAcJDg8NDA4LDAwM/9sAQwECAgIDAwMGAwMGDAgHCAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwM/8AAEQgAXQB7AwEiAAIRAQMRAf/EAB8AAAEFAQEBAQEB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");
                        //iTextSharp.text.Image stfoto = iTextSharp.text.Image.GetInstance(imageBytes);
                        //stfoto.SetAbsolutePosition(100, 500);
                        //Imagen.Alignment = Element.ALIGN_LEFT;


                        string CMS = "CAROL MORGAN SCHOOL";
                        string HSRS = "HIGH SCHOOL REPORT CARD SEMESTER 1";


                        PdfPTable HeadT = new PdfPTable(3);
                        HeadT.HorizontalAlignment = Element.ALIGN_CENTER;
                        HeadT.WidthPercentage = 100;

                        PdfPCell CMSP = new PdfPCell(new Phrase(CMS, new Font(Font.FontFamily.HELVETICA, 16, Font.BOLD, BaseColor.BLACK)));
                        CMSP.Colspan = 3;
                        CMSP.HorizontalAlignment = Element.ALIGN_CENTER;
                        CMSP.Border = 0;
                        PdfPCell HSRSP = new PdfPCell(new Phrase(HSRS, new Font(Font.FontFamily.HELVETICA, 16, Font.BOLD, BaseColor.BLACK)));
                        HSRSP.Colspan = 3;
                        HSRSP.HorizontalAlignment = Element.ALIGN_CENTER;
                        HSRSP.Border = 0;
                        PdfPCell HEAD3 = new PdfPCell(new Phrase("Semester 1 Report Card", new Font(Font.FontFamily.HELVETICA, 16, Font.BOLD, BaseColor.BLACK)));
                        HEAD3.Colspan = 3;
                        HEAD3.Border = 0;
                        HEAD3.HorizontalAlignment = Element.ALIGN_CENTER;
                        PdfPCell HEAD4 = new PdfPCell(new Phrase("Term:" + stcomm[0], new Font(Font.FontFamily.HELVETICA, 16, Font.BOLD, BaseColor.BLACK)));
                        HEAD4.Colspan = 3;
                        HEAD4.HorizontalAlignment = Element.ALIGN_CENTER;
                        HEAD4.Border = 0;
                        HeadT.AddCell(CMSP);
                        HeadT.AddCell(HSRSP);
                        HeadT.AddCell(HEAD3);
                        HeadT.AddCell(HEAD4);

                        PdfPCell stinfo = new PdfPCell(new Phrase("Student Name:" + stcomm[2], new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD, BaseColor.BLACK)));
                        stinfo.HorizontalAlignment = Element.ALIGN_LEFT;
                        stinfo.Border = 0;
                        stinfo.Colspan = 3;
                        HeadT.AddCell(stinfo);

                        PdfPCell stinfo1 = new PdfPCell(new Phrase("Grade:" + stcomm[3], new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD, BaseColor.BLACK)));
                        stinfo1.HorizontalAlignment = Element.ALIGN_LEFT;
                        stinfo1.Border = 0;
                        HeadT.AddCell(stinfo1);

                        PdfPCell stid = new PdfPCell(new Phrase("StudentID:" + stcomm[1], new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD, BaseColor.WHITE)));
                        stid.HorizontalAlignment = Element.ALIGN_LEFT;
                        stid.Border = 0;
                        HeadT.AddCell(stid);

                        PdfPCell repdate = new PdfPCell(new Phrase("Date: " + DateTime.Now.ToString("MM/dd/yyyy"), new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD, BaseColor.BLACK)));
                        repdate.HorizontalAlignment = Element.ALIGN_RIGHT;
                        repdate.Border = 0;
                        HeadT.AddCell(repdate);


                        PdfPTable GradeTable = new PdfPTable(19);
                        GradeTable.HorizontalAlignment = Element.ALIGN_CENTER;
                        GradeTable.WidthPercentage = 100;

                        PdfPCell Course = new PdfPCell(new Phrase("Course", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                        Course.HorizontalAlignment = Element.ALIGN_LEFT;
                        Course.BackgroundColor = new BaseColor(135, 0, 27);
                        Course.BorderWidth = 1F;
                        Course.Colspan = 4;
                        Course.Padding = 2f;
                        GradeTable.AddCell(Course);
                        PdfPCell Teacher = new PdfPCell(new Phrase("Teacher", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                        Teacher.HorizontalAlignment = Element.ALIGN_CENTER;
                        Teacher.BackgroundColor = new BaseColor(135, 0, 27);
                        Teacher.BorderWidth = 1F;
                        Teacher.Colspan = 2;
                        Teacher.Padding = 2f;
                        GradeTable.AddCell(Teacher);
                        PdfPCell Q1 = new PdfPCell(new Phrase("Q1", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                        Q1.HorizontalAlignment = Element.ALIGN_CENTER;
                        Q1.BackgroundColor = new BaseColor(135, 0, 27);
                        Q1.BorderWidth = 1F;
                        Q1.Padding = 2f;
                        GradeTable.AddCell(Q1);
                        PdfPCell Q2 = new PdfPCell(new Phrase("Q2", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                        Q2.HorizontalAlignment = Element.ALIGN_CENTER;
                        Q2.BackgroundColor = new BaseColor(135, 0, 27);
                        Q2.BorderWidth = 1F;
                        Q2.Padding = 2f;
                        GradeTable.AddCell(Q2);
                        PdfPCell EX1 = new PdfPCell(new Phrase("EX1", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                        EX1.HorizontalAlignment = Element.ALIGN_CENTER;
                        EX1.BackgroundColor = new BaseColor(135, 0, 27);
                        EX1.BorderWidth = 1F;
                        EX1.Padding = 2f;
                        GradeTable.AddCell(EX1);
                        PdfPCell S1 = new PdfPCell(new Phrase("S1", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                        S1.HorizontalAlignment = Element.ALIGN_CENTER;
                        S1.BackgroundColor = new BaseColor(135, 0, 27);
                        S1.BorderWidth = 1F;
                        S1.Padding = 2f;
                        GradeTable.AddCell(S1);
                        PdfPCell CommS1 = new PdfPCell(new Phrase("Comment S1", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                        CommS1.HorizontalAlignment = Element.ALIGN_LEFT;
                        CommS1.BackgroundColor = new BaseColor(135, 0, 27);
                        CommS1.BorderWidth = 1F;
                        CommS1.Colspan = 7;
                        CommS1.Padding = 2f;
                        GradeTable.AddCell(CommS1);
                        PdfPCell ABS = new PdfPCell(new Phrase("ABS", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                        ABS.HorizontalAlignment = Element.ALIGN_CENTER;
                        ABS.BackgroundColor = new BaseColor(135, 0, 27);
                        ABS.BorderWidth = 1F;
                        ABS.Padding = 2f;
                        GradeTable.AddCell(ABS);
                        PdfPCell TARD = new PdfPCell(new Phrase("TAR", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                        TARD.HorizontalAlignment = Element.ALIGN_CENTER;
                        TARD.BackgroundColor = new BaseColor(135, 0, 27);
                        TARD.BorderWidth = 1F;
                        TARD.Padding = 2f;
                        GradeTable.AddCell(TARD);



                        PdfPCell CO;
                        PdfPCell TE;
                        PdfPCell QU1;
                        PdfPCell QU2;
                        PdfPCell EXE1;
                        PdfPCell SE1;
                        PdfPCell COM1;
                        PdfPCell ABS1;
                        PdfPCell TAR1;

                        for (int i = 0; i < stTable.Length - 1; i++)
                        {
                            var nfila = stTable[i].Split('|');

                            CO = new PdfPCell(new Phrase(nfila[0], new Font(Font.FontFamily.HELVETICA, 9.0F, Font.NORMAL, BaseColor.BLACK)));
                            CO.HorizontalAlignment = Element.ALIGN_LEFT;
                            CO.BorderWidthLeft = 1.0F;
                            CO.Colspan = 4;
                            GradeTable.AddCell(CO);
                            TE = new PdfPCell(new Phrase(nfila[1], new Font(Font.FontFamily.HELVETICA, 9.0F, Font.NORMAL, BaseColor.BLACK)));
                            TE.HorizontalAlignment = Element.ALIGN_CENTER;
                            TE.BorderWidthLeft = 1.0F;
                            TE.Colspan = 2;
                            GradeTable.AddCell(TE);
                            QU1 = new PdfPCell(new Phrase(nfila[2], new Font(Font.FontFamily.HELVETICA, 9.0F, Font.NORMAL, BaseColor.BLACK)));
                            QU1.HorizontalAlignment = Element.ALIGN_CENTER;
                            QU1.BorderWidthLeft = 1.0F;
                            GradeTable.AddCell(QU1);
                            QU2 = new PdfPCell(new Phrase(nfila[3], new Font(Font.FontFamily.HELVETICA, 9.0F, Font.NORMAL, BaseColor.BLACK)));
                            QU2.HorizontalAlignment = Element.ALIGN_CENTER;
                            QU2.BorderWidthLeft = 1.0F;
                            GradeTable.AddCell(QU2);
                            EXE1 = new PdfPCell(new Phrase(nfila[4], new Font(Font.FontFamily.HELVETICA, 9.0F, Font.NORMAL, BaseColor.BLACK)));
                            EXE1.HorizontalAlignment = Element.ALIGN_CENTER;
                            EXE1.BorderWidthLeft = 1.0F;
                            GradeTable.AddCell(EXE1);
                            SE1 = new PdfPCell(new Phrase(nfila[5], new Font(Font.FontFamily.HELVETICA, 9.0F, Font.NORMAL, BaseColor.BLACK)));
                            SE1.HorizontalAlignment = Element.ALIGN_CENTER;
                            SE1.BorderWidthLeft = 1.0F;
                            GradeTable.AddCell(SE1);
                            COM1 = new PdfPCell(new Phrase(nfila[6], new Font(Font.FontFamily.HELVETICA, 9.0F, Font.NORMAL, BaseColor.BLACK)));
                            COM1.HorizontalAlignment = Element.ALIGN_LEFT;
                            COM1.BorderWidthLeft = 1.0F;
                            COM1.Colspan = 7;
                            GradeTable.AddCell(COM1);
                            ABS1 = new PdfPCell(new Phrase(nfila[7], new Font(Font.FontFamily.HELVETICA, 9.0F, Font.NORMAL, BaseColor.BLACK)));
                            ABS1.HorizontalAlignment = Element.ALIGN_CENTER;
                            ABS1.BorderWidthLeft = 1.0F;
                            GradeTable.AddCell(ABS1);
                            TAR1 = new PdfPCell(new Phrase(nfila[8], new Font(Font.FontFamily.HELVETICA, 9.0F, Font.NORMAL, BaseColor.BLACK)));
                            TAR1.HorizontalAlignment = Element.ALIGN_CENTER;
                            TAR1.BorderWidthLeft = 1.0F;
                            GradeTable.AddCell(TAR1);


                        }

                        PdfPTable FOOT = new PdfPTable(8);
                        FOOT.HorizontalAlignment = Element.ALIGN_LEFT;
                        FOOT.WidthPercentage = 100;

                        PdfPCell Semgpa = new PdfPCell(new Phrase("Semester And Cumulative GPA", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.BLACK)));
                        Semgpa.HorizontalAlignment = Element.ALIGN_LEFT;
                        Semgpa.Colspan = 4;
                        Semgpa.Border = 0;
                        FOOT.AddCell(Semgpa);

                        PdfPCell Comms = new PdfPCell(new Phrase("Coomunity Service Hours", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.BLACK)));
                        Comms.HorizontalAlignment = Element.ALIGN_LEFT;
                        Comms.Colspan = 4;
                        Comms.Border = 0;
                        FOOT.AddCell(Comms);

                        PdfPCell Sem1 = new PdfPCell(new Phrase("Semester 1 GPA:" + stS1GPA[1], new Font(Font.FontFamily.HELVETICA, 11.0F, Font.NORMAL, BaseColor.BLACK)));
                        Sem1.HorizontalAlignment = Element.ALIGN_LEFT;
                        Sem1.Colspan = 4;
                        Sem1.Border = 0;
                        FOOT.AddCell(Sem1);

                        PdfPCell Curr = new PdfPCell(new Phrase("Current Year Hours:" + stcomm[5], new Font(Font.FontFamily.HELVETICA, 11.0F, Font.NORMAL, BaseColor.BLACK)));
                        Curr.HorizontalAlignment = Element.ALIGN_LEFT;
                        Curr.Colspan = 4;
                        Curr.Border = 0;
                        FOOT.AddCell(Curr);

                        PdfPCell Cumul = new PdfPCell(new Phrase("Cumulative GPA:" + stCUGPA[1], new Font(Font.FontFamily.HELVETICA, 11.0F, Font.NORMAL, BaseColor.BLACK)));
                        Cumul.HorizontalAlignment = Element.ALIGN_LEFT;
                        Cumul.Colspan = 4;
                        Cumul.Border = 0;
                        FOOT.AddCell(Cumul);

                        PdfPCell InH = new PdfPCell(new Phrase("In School Hours:" + stcomm[6], new Font(Font.FontFamily.HELVETICA, 11.0F, Font.NORMAL, BaseColor.BLACK)));
                        InH.HorizontalAlignment = Element.ALIGN_LEFT;
                        InH.Colspan = 4;
                        InH.Border = 0;
                        FOOT.AddCell(InH);

                        PdfPCell spa = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                        spa.Colspan = 4;
                        spa.Border = 0;
                        FOOT.AddCell(spa);

                        PdfPCell outH = new PdfPCell(new Phrase("OutReach Hours:" + stcomm[7], new Font(Font.FontFamily.HELVETICA, 11.0F, Font.NORMAL, BaseColor.BLACK)));
                        outH.HorizontalAlignment = Element.ALIGN_LEFT;
                        outH.Colspan = 4;
                        outH.Border = 0;
                        FOOT.AddCell(outH);

                        PdfPCell HRD = new PdfPCell(new Phrase("Honor Roll Description", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.BLACK)));
                        HRD.HorizontalAlignment = Element.ALIGN_LEFT;

                        HRD.Colspan = 4;
                        HRD.Border = 0;
                        FOOT.AddCell(HRD);


                        PdfPCell Total = new PdfPCell(new Phrase("Total (Grade 9-12) Hours:" + stcomm[8], new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.BLACK)));
                        Total.HorizontalAlignment = Element.ALIGN_LEFT;
                        Total.Colspan = 4;
                        Total.Border = 0;
                        FOOT.AddCell(Total);

                        PdfPCell HRDe = new PdfPCell(new Phrase("Principal -GPA 4.50 - No Grade Below 90 (AP85).", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.NORMAL, BaseColor.BLACK)));
                        HRDe.HorizontalAlignment = Element.ALIGN_LEFT;
                        HRDe.Colspan = 4;
                        HRDe.Border = 0;
                        FOOT.AddCell(HRDe);

                        PdfPCell hr = new PdfPCell(new Phrase("15 hours requiered each year (Minimum of 10 Outreach) 60 hours minimum requiered in grades 9 to 12.", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.NORMAL, BaseColor.BLACK)));
                        hr.HorizontalAlignment = Element.ALIGN_LEFT;
                        hr.Colspan = 4;
                        hr.Border = 0;
                        hr.Rowspan = 2;
                        FOOT.AddCell(hr);

                        PdfPCell Hih = new PdfPCell(new Phrase("High Honors - GPA 4.00 - No Grade Below 85 (AP 80).", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.NORMAL, BaseColor.BLACK)));
                        Hih.HorizontalAlignment = Element.ALIGN_LEFT;
                        Hih.Colspan = 8;
                        Hih.Border = 0;
                        FOOT.AddCell(Hih);
                        PdfPCell Hon = new PdfPCell(new Phrase("Honors - GPA 3.50 - No Grade Below 80 (AP 75).", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.NORMAL, BaseColor.BLACK)));
                        Hon.HorizontalAlignment = Element.ALIGN_LEFT;
                        Hon.Colspan = 8;
                        Hon.Border = 0;
                        FOOT.AddCell(Hon);

                        PdfPCell spa3 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.NORMAL, BaseColor.BLACK)));
                        spa3.Colspan = 8;
                        spa3.Border = 0;
                        FOOT.AddCell(spa3);

                        PdfPCell IFY = new PdfPCell(new Phrase("If you have any questions regarding your child's Progress Report, Please contact High School Office: 809-947-1033.", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.NORMAL, BaseColor.BLACK)));
                        IFY.HorizontalAlignment = Element.ALIGN_LEFT;
                        IFY.Colspan = 4;
                        IFY.Border = 0;
                        FOOT.AddCell(IFY);



                        PdfPCell not = new PdfPCell(new Phrase("Note: Absences and tardiness displayed in the report correspond to first semester of the year.", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.NORMAL, BaseColor.BLACK)));
                        not.HorizontalAlignment = Element.ALIGN_LEFT;
                        not.Colspan = 4;
                        not.Border = 0;
                        FOOT.AddCell(not);

                        documento.Add(Imagen);
                        //documento.Add(stfoto);
                        documento.Add(HeadT);
                        documento.Add(GradeTable);
                        documento.Add(FOOT);

                        //Process prc = new System.Diagnostics.Process();
                        //prc.StartInfo.FileName = fileName;
                        //prc.Start();
                    }
                    else
                    {
                        con.Close();
                      
                    }
                        documento.NewPage();
                    }
                    
                    con.Close();
                   
                }
                else
                {
                    if (con.State == ConnectionState.Closed)
                    {
                        con.Open();
                    }
                    string stdata = string.Empty;
                    string S1GPA = string.Empty;
                    string CUGPA = string.Empty;
                    string datat = string.Empty;

                    sql = "WITH main_query AS(SELECT DISTINCT S.STUDENT_NUMBER, SG.COURSE_NAME, SG.STORECODE, SG.GRADE, SG.TEACHER_NAME,";
                    sql += " TO_CHAR(SG.COMMENT_VALUE) AS COMMENTS, SG.ABSENCES, SG.TARDIES FROM   STOREDGRADES SG";
                    sql += " LEFT JOIN STUDENTS S ON SG.STUDENTID = S.ID";
                    sql += " WHERE SG.TERMID IN(2700, 2701)  AND STORECODE IN('Q1', 'Q2', 'E1', 'S1') AND S.STUDENT_NUMBER =" + stnum + "";
                    sql += " )";
                    sql += " SELECT DISTINCT COURSE_NAME, TEACHER_NAME";
                    sql += " ,(SELECT  y.GRADE FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.STORECODE = 'Q1') Q1";
                    sql += " ,(SELECT  y.GRADE FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.STORECODE = 'Q2') Q2";
                    sql += " ,(SELECT  y.GRADE FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.STORECODE = 'E1') E1";
                    sql += " ,(SELECT  y.GRADE FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.STORECODE = 'S1') S1";
                    sql += " ,(SELECT  y.COMMENTS FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.STORECODE = 'S1') Comments";
                    sql += " ,(SELECT  y.ABSENCES FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.STORECODE = 'S1') ABS1";
                    sql += " ,(SELECT  y.TARDIES FROM main_query y WHERE y.COURSE_NAME = M.COURSE_NAME and y.STORECODE = 'S1') TARDI";
                    sql += "  FROM main_query M";
                    sql += " ORDER BY COURSE_NAME";

                    OracleCommand cmd = new OracleCommand(sql, con);
                    OracleDataReader odr = cmd.ExecuteReader();
                    while (odr.Read())
                    {
                        datat += odr["COURSE_NAME"].ToString() + '|';
                        datat += odr["TEACHER_NAME"].ToString() + '|';
                        datat += odr["Q1"].ToString() + '|';
                        datat += odr["Q2"].ToString() + '|';
                        datat += odr["E1"].ToString() + '|';
                        datat += odr["S1"].ToString() + '|';
                        datat += odr["Comments"].ToString() + '|';
                        datat += odr["ABS1"].ToString() + '|';
                        datat += odr["TARDI"].ToString() + '^';

                    }
                    if (datat != "")
                    {
                        // Close and Dispose OracleConnection object

                        sql = "SELECT DISTINCT T.ABBREVIATION,s.student_number,s.lastfirst,s.grade_level,S.studentpict_guid,CO.HS_SERVICE_HOURS_CURRENT,CO.HS_SERVICE_HOURS_CURRENT_IN,CO.HS_SERVICE_HOURS_CURRENT_OUTRE,CO.HS_TOTAL_SERVICE_HOURS FROM STUDENTS S";
                        sql += " LEFT JOIN U_COMMUNITYSERVICE CO ON S.DCID=CO.STUDENTSDCID";
                        sql += " LEFT JOIN STOREDGRADES SG ON S.ID=SG.STUDENTID";
                        sql += " LEFT JOIN TERMS T ON SG.TERMID=T.ID";
                        sql += " WHERE SG.TERMID IN (2700) AND S.STUDENT_NUMBER =" + stnum + "";

                        OracleCommand cmd1 = new OracleCommand(sql, con);
                        OracleDataReader odr1 = cmd1.ExecuteReader();
                        while (odr1.Read())
                        {
                            stdata += odr1["ABBREVIATION"].ToString() + '|';
                            stdata += odr1["student_number"].ToString() + '|';
                            stdata += odr1["lastfirst"].ToString() + '|';
                            stdata += odr1["grade_level"].ToString() + '|';
                            stdata += odr1["studentpict_guid"].ToString() + '|';
                            stdata += odr1["HS_SERVICE_HOURS_CURRENT"].ToString() + '|';
                            stdata += odr1["HS_SERVICE_HOURS_CURRENT_IN"].ToString() + '|';
                            stdata += odr1["HS_SERVICE_HOURS_CURRENT_OUTRE"].ToString() + '|';
                            stdata += odr1["HS_TOTAL_SERVICE_HOURS"].ToString() + '|';

                        }

                        sql = "SELECT S.STUDENT_NUMBER, ROUND(SUM(sg.gpa_points)/COUNT(sg.gpa_points),3) AS GPA FROM STOREDGRADES SG";
                        sql += " LEFT JOIN STUDENTS S ON SG.STUDENTID=S.ID";
                        sql += " WHERE SG.TERMID IN (2700,2702,2701) AND S.STUDENT_NUMBER =" + stnum + "";
                        sql += " AND SG.STORECODE IN ('S1') AND SG.GPA_POINTS<>0 ";
                        sql += " GROUP BY S.STUDENT_NUMBER";


                        OracleCommand cmd2 = new OracleCommand(sql, con);
                        OracleDataReader odr2 = cmd2.ExecuteReader();
                        while (odr2.Read())
                        {
                            S1GPA += odr2["STUDENT_NUMBER"].ToString() + '|';
                            S1GPA += odr2["GPA"].ToString() + '|';

                        }

                        sql = "SELECT S.STUDENT_NUMBER, ROUND(SUM(sg.gpa_points)/COUNT(sg.gpa_points),3) AS GPA FROM STOREDGRADES SG";
                        sql += " LEFT JOIN STUDENTS S ON SG.STUDENTID=S.ID";
                        sql += " WHERE SG.TERMID IN (2700,2702,2701,2600,2602,2601,2500,2502,2501) AND S.STUDENT_NUMBER =" + stnum + "";
                        sql += " AND SG.STORECODE IN ('S1','S2') AND SG.GPA_POINTS<>0 ";
                        sql += " GROUP BY S.STUDENT_NUMBER";


                        OracleCommand cmd3 = new OracleCommand(sql, con);
                        OracleDataReader odr3 = cmd3.ExecuteReader();
                        while (odr3.Read())
                        {
                            CUGPA += odr3["STUDENT_NUMBER"].ToString() + '|';
                            CUGPA += odr3["GPA"].ToString() + '|';

                        }

                        // Close and Dispose OracleConnection object
                       

                        var stTable = datat.Split('^');
                        var stcomm = stdata.Split('|');
                        var stS1GPA = S1GPA.Split('|');
                        var stCUGPA = CUGPA.Split('|');

                        fname = "HS_ReportCardS1_" + DateTime.Now.DayOfYear + DateTime.Now.Month + DateTime.Now.Day + DateTime.Now.Millisecond + ".pdf";
                        fileName = HttpContext.Current.Server.MapPath("~/RepoFiles/" + fname);

                        PdfWriter.GetInstance(documento, new FileStream(fileName, FileMode.Create));

                        documento.Open();

                        iTextSharp.text.Image Imagen = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("~/img/cmslogo.jpg"));
                        Imagen.SetAbsolutePosition(-3, 520);
                        Imagen.ScalePercent(2f);

                        //byte[] imageBytes = Convert.FromBase64String(@"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LyIgeGFwUmlnaHRzOk1hcmtlZD0iVHJ1ZSIgeGFwUmlnaHRzOldlYlN0YXRlbWVudD0iaHR0cDovL3Byby5jb3JiaXMuY29tL3NlYXJjaC9zZWFyY2hyZXN1bHRzLmFzcD90eHQ9NDItMTU1NjQ5NzgmYW1wO29wZW5JbWFnZT00Mi0xNTU2NDk3OCI+DQoJCQk8ZGM6cmlnaHRzPg0KCQkJCTxyZGY6QWx0Pg0KCQkJCQk8cmRmOmxpIHhtbDpsYW5nPSJ4LWRlZmF1bHQiPsKpIENvcmJpcy4gIEFsbCBSaWdodHMgUmVzZXJ2ZWQuPC9yZGY6bGk+DQoJCQkJPC9yZGY6QWx0Pg0KCQkJPC9kYzpyaWdodHM+DQoJCQk8ZGM6Y3JlYXRvcj48cmRmOlNlcSB4bWxuczpyZGY9Imh0dHA6Ly93d3cudzMub3JnLzE5OTkvMDIvMjItcmRmLXN5bnRheC1ucyMiPjxyZGY6bGk+Q29yYmlzPC9yZGY6bGk+PC9yZGY6U2VxPg0KCQkJPC9kYzpjcmVhdG9yPjwvcmRmOkRlc2NyaXB0aW9uPg0KCQk8cmRmOkRlc2NyaXB0aW9uIHhtbG5zOnhtcD0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wLyI+PHhtcDpSYXRpbmc+NDwveG1wOlJhdGluZz48eG1wOkNyZWF0ZURhdGU+MjAwOC0wMi0xMVQxOTozMjo0My4xNzNaPC94bXA6Q3JlYXRlRGF0ZT48L3JkZjpEZXNjcmlwdGlvbj48cmRmOkRlc2NyaXB0aW9uIHhtbG5zOk1pY3Jvc29mdFBob3RvPSJodHRwOi8vbnMubWljcm9zb2Z0LmNvbS9waG90by8xLjAvIj48TWljcm9zb2Z0UGhvdG86UmF0aW5nPjYzPC9NaWNyb3NvZnRQaG90bzpSYXRpbmc+PC9yZGY6RGVzY3JpcHRpb24+PC9yZGY6UkRGPg0KPC94OnhtcG1ldGE+DQogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICA8P3hwYWNrZXQgZW5kPSd3Jz8+/9sAQwACAQECAQECAgICAgICAgMFAwMDAwMGBAQDBQcGBwcHBgcHCAkLCQgICggHBwoNCgoLDAwMDAcJDg8NDA4LDAwM/9sAQwECAgIDAwMGAwMGDAgHCAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwM/8AAEQgAXQB7AwEiAAIRAQMRAf/EAB8AAAEFAQEBAQEB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");
                        //iTextSharp.text.Image stfoto = iTextSharp.text.Image.GetInstance(imageBytes);
                        //stfoto.SetAbsolutePosition(100, 500);
                        //Imagen.Alignment = Element.ALIGN_LEFT;


                        string CMS = "CAROL MORGAN SCHOOL";
                        string HSRS = "HIGH SCHOOL REPORT CARD SEMESTER 1";


                        PdfPTable HeadT = new PdfPTable(3);
                        HeadT.HorizontalAlignment = Element.ALIGN_CENTER;
                        HeadT.WidthPercentage = 100;

                        PdfPCell CMSP = new PdfPCell(new Phrase(CMS, new Font(Font.FontFamily.HELVETICA, 16, Font.BOLD, BaseColor.BLACK)));
                        CMSP.Colspan = 3;
                        CMSP.HorizontalAlignment = Element.ALIGN_CENTER;
                        CMSP.Border = 0;
                        PdfPCell HSRSP = new PdfPCell(new Phrase(HSRS, new Font(Font.FontFamily.HELVETICA, 16, Font.BOLD, BaseColor.BLACK)));
                        HSRSP.Colspan = 3;
                        HSRSP.HorizontalAlignment = Element.ALIGN_CENTER;
                        HSRSP.Border = 0;
                        PdfPCell HEAD3 = new PdfPCell(new Phrase("Semester 1 Report Card", new Font(Font.FontFamily.HELVETICA, 16, Font.BOLD, BaseColor.BLACK)));
                        HEAD3.Colspan = 3;
                        HEAD3.Border = 0;
                        HEAD3.HorizontalAlignment = Element.ALIGN_CENTER;
                        PdfPCell HEAD4 = new PdfPCell(new Phrase("Term:" + stcomm[0], new Font(Font.FontFamily.HELVETICA, 16, Font.BOLD, BaseColor.BLACK)));
                        HEAD4.Colspan = 3;
                        HEAD4.HorizontalAlignment = Element.ALIGN_CENTER;
                        HEAD4.Border = 0;
                        HeadT.AddCell(CMSP);
                        HeadT.AddCell(HSRSP);
                        HeadT.AddCell(HEAD3);
                        HeadT.AddCell(HEAD4);

                        PdfPCell stinfo = new PdfPCell(new Phrase("Student Name:" + stcomm[2], new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD, BaseColor.BLACK)));
                        stinfo.HorizontalAlignment = Element.ALIGN_LEFT;
                        stinfo.Border = 0;
                        stinfo.Colspan = 3;
                        HeadT.AddCell(stinfo);

                        PdfPCell stinfo1 = new PdfPCell(new Phrase("Grade:" + stcomm[3], new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD, BaseColor.BLACK)));
                        stinfo1.HorizontalAlignment = Element.ALIGN_LEFT;
                        stinfo1.Border = 0;
                        HeadT.AddCell(stinfo1);

                        PdfPCell stid = new PdfPCell(new Phrase("StudentID:" + stcomm[1], new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD, BaseColor.WHITE)));
                        stid.HorizontalAlignment = Element.ALIGN_LEFT;
                        stid.Border = 0;
                        HeadT.AddCell(stid);

                        PdfPCell repdate = new PdfPCell(new Phrase("Date: " + DateTime.Now.ToString("MM/dd/yyyy"), new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD, BaseColor.BLACK)));
                        repdate.HorizontalAlignment = Element.ALIGN_RIGHT;
                        repdate.Border = 0;
                        HeadT.AddCell(repdate);


                        PdfPTable GradeTable = new PdfPTable(19);
                        GradeTable.HorizontalAlignment = Element.ALIGN_CENTER;
                        GradeTable.WidthPercentage = 100;

                        PdfPCell Course = new PdfPCell(new Phrase("Course", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                        Course.HorizontalAlignment = Element.ALIGN_LEFT;
                        Course.BackgroundColor = new BaseColor(135, 0, 27);
                        Course.BorderWidth = 1F;
                        Course.Colspan = 4;
                        Course.Padding = 2f;
                        GradeTable.AddCell(Course);
                        PdfPCell Teacher = new PdfPCell(new Phrase("Teacher", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                        Teacher.HorizontalAlignment = Element.ALIGN_CENTER;
                        Teacher.BackgroundColor = new BaseColor(135, 0, 27);
                        Teacher.BorderWidth = 1F;
                        Teacher.Colspan = 2;
                        Teacher.Padding = 2f;
                        GradeTable.AddCell(Teacher);
                        PdfPCell Q1 = new PdfPCell(new Phrase("Q1", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                        Q1.HorizontalAlignment = Element.ALIGN_CENTER;
                        Q1.BackgroundColor = new BaseColor(135, 0, 27);
                        Q1.BorderWidth = 1F;
                        Q1.Padding = 2f;
                        GradeTable.AddCell(Q1);
                        PdfPCell Q2 = new PdfPCell(new Phrase("Q2", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                        Q2.HorizontalAlignment = Element.ALIGN_CENTER;
                        Q2.BackgroundColor = new BaseColor(135, 0, 27);
                        Q2.BorderWidth = 1F;
                        Q2.Padding = 2f;
                        GradeTable.AddCell(Q2);
                        PdfPCell EX1 = new PdfPCell(new Phrase("EX1", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                        EX1.HorizontalAlignment = Element.ALIGN_CENTER;
                        EX1.BackgroundColor = new BaseColor(135, 0, 27);
                        EX1.BorderWidth = 1F;
                        EX1.Padding = 2f;
                        GradeTable.AddCell(EX1);
                        PdfPCell S1 = new PdfPCell(new Phrase("S1", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                        S1.HorizontalAlignment = Element.ALIGN_CENTER;
                        S1.BackgroundColor = new BaseColor(135, 0, 27);
                        S1.BorderWidth = 1F;
                        S1.Padding = 2f;
                        GradeTable.AddCell(S1);
                        PdfPCell CommS1 = new PdfPCell(new Phrase("Comment S1", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                        CommS1.HorizontalAlignment = Element.ALIGN_LEFT;
                        CommS1.BackgroundColor = new BaseColor(135, 0, 27);
                        CommS1.BorderWidth = 1F;
                        CommS1.Colspan = 7;
                        CommS1.Padding = 2f;
                        GradeTable.AddCell(CommS1);
                        PdfPCell ABS = new PdfPCell(new Phrase("ABS", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                        ABS.HorizontalAlignment = Element.ALIGN_CENTER;
                        ABS.BackgroundColor = new BaseColor(135, 0, 27);
                        ABS.BorderWidth = 1F;
                        ABS.Padding = 2f;
                        GradeTable.AddCell(ABS);
                        PdfPCell TARD = new PdfPCell(new Phrase("TAR", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.BOLD, BaseColor.WHITE)));
                        TARD.HorizontalAlignment = Element.ALIGN_CENTER;
                        TARD.BackgroundColor = new BaseColor(135, 0, 27);
                        TARD.BorderWidth = 1F;
                        TARD.Padding = 2f;
                        GradeTable.AddCell(TARD);



                        PdfPCell CO;
                        PdfPCell TE;
                        PdfPCell QU1;
                        PdfPCell QU2;
                        PdfPCell EXE1;
                        PdfPCell SE1;
                        PdfPCell COM1;
                        PdfPCell ABS1;
                        PdfPCell TAR1;

                        for (int i = 0; i < stTable.Length - 1; i++)
                        {
                            var nfila = stTable[i].Split('|');

                            CO = new PdfPCell(new Phrase(nfila[0], new Font(Font.FontFamily.HELVETICA, 9.0F, Font.NORMAL, BaseColor.BLACK)));
                            CO.HorizontalAlignment = Element.ALIGN_LEFT;
                            CO.BorderWidthLeft = 1.0F;
                            CO.Colspan = 4;
                            GradeTable.AddCell(CO);
                            TE = new PdfPCell(new Phrase(nfila[1], new Font(Font.FontFamily.HELVETICA, 9.0F, Font.NORMAL, BaseColor.BLACK)));
                            TE.HorizontalAlignment = Element.ALIGN_CENTER;
                            TE.BorderWidthLeft = 1.0F;
                            TE.Colspan = 2;
                            GradeTable.AddCell(TE);
                            QU1 = new PdfPCell(new Phrase(nfila[2], new Font(Font.FontFamily.HELVETICA, 9.0F, Font.NORMAL, BaseColor.BLACK)));
                            QU1.HorizontalAlignment = Element.ALIGN_CENTER;
                            QU1.BorderWidthLeft = 1.0F;
                            GradeTable.AddCell(QU1);
                            QU2 = new PdfPCell(new Phrase(nfila[3], new Font(Font.FontFamily.HELVETICA, 9.0F, Font.NORMAL, BaseColor.BLACK)));
                            QU2.HorizontalAlignment = Element.ALIGN_CENTER;
                            QU2.BorderWidthLeft = 1.0F;
                            GradeTable.AddCell(QU2);
                            EXE1 = new PdfPCell(new Phrase(nfila[4], new Font(Font.FontFamily.HELVETICA, 9.0F, Font.NORMAL, BaseColor.BLACK)));
                            EXE1.HorizontalAlignment = Element.ALIGN_CENTER;
                            EXE1.BorderWidthLeft = 1.0F;
                            GradeTable.AddCell(EXE1);
                            SE1 = new PdfPCell(new Phrase(nfila[5], new Font(Font.FontFamily.HELVETICA, 9.0F, Font.NORMAL, BaseColor.BLACK)));
                            SE1.HorizontalAlignment = Element.ALIGN_CENTER;
                            SE1.BorderWidthLeft = 1.0F;
                            GradeTable.AddCell(SE1);
                            COM1 = new PdfPCell(new Phrase(nfila[6], new Font(Font.FontFamily.HELVETICA, 9.0F, Font.NORMAL, BaseColor.BLACK)));
                            COM1.HorizontalAlignment = Element.ALIGN_LEFT;
                            COM1.BorderWidthLeft = 1.0F;
                            COM1.Colspan = 7;
                            GradeTable.AddCell(COM1);
                            ABS1 = new PdfPCell(new Phrase(nfila[7], new Font(Font.FontFamily.HELVETICA, 9.0F, Font.NORMAL, BaseColor.BLACK)));
                            ABS1.HorizontalAlignment = Element.ALIGN_CENTER;
                            ABS1.BorderWidthLeft = 1.0F;
                            GradeTable.AddCell(ABS1);
                            TAR1 = new PdfPCell(new Phrase(nfila[8], new Font(Font.FontFamily.HELVETICA, 9.0F, Font.NORMAL, BaseColor.BLACK)));
                            TAR1.HorizontalAlignment = Element.ALIGN_CENTER;
                            TAR1.BorderWidthLeft = 1.0F;
                            GradeTable.AddCell(TAR1);


                        }

                        PdfPTable FOOT = new PdfPTable(8);
                        FOOT.HorizontalAlignment = Element.ALIGN_LEFT;
                        FOOT.WidthPercentage = 100;

                        PdfPCell Semgpa = new PdfPCell(new Phrase("Semester And Cumulative GPA", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.BLACK)));
                        Semgpa.HorizontalAlignment = Element.ALIGN_LEFT;
                        Semgpa.Colspan = 4;
                        Semgpa.Border = 0;
                        FOOT.AddCell(Semgpa);

                        PdfPCell Comms = new PdfPCell(new Phrase("Coomunity Service Hours", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.BLACK)));
                        Comms.HorizontalAlignment = Element.ALIGN_LEFT;
                        Comms.Colspan = 4;
                        Comms.Border = 0;
                        FOOT.AddCell(Comms);

                        PdfPCell Sem1 = new PdfPCell(new Phrase("Semester 1 GPA:" + stS1GPA[1], new Font(Font.FontFamily.HELVETICA, 11.0F, Font.NORMAL, BaseColor.BLACK)));
                        Sem1.HorizontalAlignment = Element.ALIGN_LEFT;
                        Sem1.Colspan = 4;
                        Sem1.Border = 0;
                        FOOT.AddCell(Sem1);

                        PdfPCell Curr = new PdfPCell(new Phrase("Current Year Hours:" + stcomm[5], new Font(Font.FontFamily.HELVETICA, 11.0F, Font.NORMAL, BaseColor.BLACK)));
                        Curr.HorizontalAlignment = Element.ALIGN_LEFT;
                        Curr.Colspan = 4;
                        Curr.Border = 0;
                        FOOT.AddCell(Curr);

                        PdfPCell Cumul = new PdfPCell(new Phrase("Cumulative GPA:" + stCUGPA[1], new Font(Font.FontFamily.HELVETICA, 11.0F, Font.NORMAL, BaseColor.BLACK)));
                        Cumul.HorizontalAlignment = Element.ALIGN_LEFT;
                        Cumul.Colspan = 4;
                        Cumul.Border = 0;
                        FOOT.AddCell(Cumul);

                        PdfPCell InH = new PdfPCell(new Phrase("In School Hours:" + stcomm[6], new Font(Font.FontFamily.HELVETICA, 11.0F, Font.NORMAL, BaseColor.BLACK)));
                        InH.HorizontalAlignment = Element.ALIGN_LEFT;
                        InH.Colspan = 4;
                        InH.Border = 0;
                        FOOT.AddCell(InH);

                        PdfPCell spa = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 10.0F, Font.NORMAL, BaseColor.BLACK)));
                        spa.Colspan = 4;
                        spa.Border = 0;
                        FOOT.AddCell(spa);

                        PdfPCell outH = new PdfPCell(new Phrase("OutReach Hours:" + stcomm[7], new Font(Font.FontFamily.HELVETICA, 11.0F, Font.NORMAL, BaseColor.BLACK)));
                        outH.HorizontalAlignment = Element.ALIGN_LEFT;
                        outH.Colspan = 4;
                        outH.Border = 0;
                        FOOT.AddCell(outH);

                        PdfPCell HRD = new PdfPCell(new Phrase("Honor Roll Description", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.BLACK)));
                        HRD.HorizontalAlignment = Element.ALIGN_LEFT;

                        HRD.Colspan = 4;
                        HRD.Border = 0;
                        FOOT.AddCell(HRD);


                        PdfPCell Total = new PdfPCell(new Phrase("Total (Grade 9-12) Hours:" + stcomm[8], new Font(Font.FontFamily.HELVETICA, 11.0F, Font.BOLD, BaseColor.BLACK)));
                        Total.HorizontalAlignment = Element.ALIGN_LEFT;
                        Total.Colspan = 4;
                        Total.Border = 0;
                        FOOT.AddCell(Total);

                        PdfPCell HRDe = new PdfPCell(new Phrase("Principal -GPA 4.50 - No Grade Below 90 (AP85).", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.NORMAL, BaseColor.BLACK)));
                        HRDe.HorizontalAlignment = Element.ALIGN_LEFT;
                        HRDe.Colspan = 4;
                        HRDe.Border = 0;
                        FOOT.AddCell(HRDe);

                        PdfPCell hr = new PdfPCell(new Phrase("15 hours requiered each year (Minimum of 10 Outreach) 60 hours minimum requiered in grades 9 to 12.", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.NORMAL, BaseColor.BLACK)));
                        hr.HorizontalAlignment = Element.ALIGN_LEFT;
                        hr.Colspan = 4;
                        hr.Border = 0;
                        hr.Rowspan = 2;
                        FOOT.AddCell(hr);

                        PdfPCell Hih = new PdfPCell(new Phrase("High Honors - GPA 4.00 - No Grade Below 85 (AP 80).", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.NORMAL, BaseColor.BLACK)));
                        Hih.HorizontalAlignment = Element.ALIGN_LEFT;
                        Hih.Colspan = 8;
                        Hih.Border = 0;
                        FOOT.AddCell(Hih);
                        PdfPCell Hon = new PdfPCell(new Phrase("Honors - GPA 3.50 - No Grade Below 80 (AP 75).", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.NORMAL, BaseColor.BLACK)));
                        Hon.HorizontalAlignment = Element.ALIGN_LEFT;
                        Hon.Colspan = 8;
                        Hon.Border = 0;
                        FOOT.AddCell(Hon);

                        PdfPCell spa3 = new PdfPCell(new Phrase(" ", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.NORMAL, BaseColor.BLACK)));
                        spa3.Colspan = 8;
                        spa3.Border = 0;
                        FOOT.AddCell(spa3);

                        PdfPCell IFY = new PdfPCell(new Phrase("If you have any questions regarding your child's Progress Report, Please contact High School Office: 809-947-1033.", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.NORMAL, BaseColor.BLACK)));
                        IFY.HorizontalAlignment = Element.ALIGN_LEFT;
                        IFY.Colspan = 4;
                        IFY.Border = 0;
                        FOOT.AddCell(IFY);



                        PdfPCell not = new PdfPCell(new Phrase("Note: Absences and tardiness displayed in the report correspond to first semester of the year.", new Font(Font.FontFamily.HELVETICA, 11.0F, Font.NORMAL, BaseColor.BLACK)));
                        not.HorizontalAlignment = Element.ALIGN_LEFT;
                        not.Colspan = 4;
                        not.Border = 0;
                        FOOT.AddCell(not);

                        //documento.Add(stfoto);
                        documento.Add(HeadT);
                        documento.Add(GradeTable);
                        documento.Add(FOOT);


                      

                        //Process prc = new System.Diagnostics.Process();
                        //prc.StartInfo.FileName = fileName;
                        //prc.Start();
                    }
                    else
                    {
                        con.Close();
                        fname = "";
                    }
                    con.Close();
                    
                }

                documento.Close();
                con.Dispose();
            }
            catch (Exception ex)
            {
                throw;
            }
            return fname;
            }
        }


}