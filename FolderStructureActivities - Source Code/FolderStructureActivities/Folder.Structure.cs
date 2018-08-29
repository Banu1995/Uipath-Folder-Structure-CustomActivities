using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FolderStructureActivities
{
    public class FolderStructure : CodeActivity
    {
        [Category("Input")]
        public InArgument<String> FolderPath1 { get; set; }

        [Category("Input")]
        public InArgument<String> FolderPath2 { get; set; }

        [Category("Input")]
        public InArgument<String> SaveAsPath { get; set; }

        [Category("Output")]
        public OutArgument<String> Output { get; set; }

        [STAThread]
        protected override void Execute(CodeActivityContext context)
        {
            string extract = FolderPath1.Get(context);
            string extract1 = FolderPath2.Get(context);
            string outputPath = SaveAsPath.Get(context);

            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook excelworkBook;
            Microsoft.Office.Interop.Excel.Worksheet excelSheet1, excelSheet2, excelSheet3, excelSheet4;
            Microsoft.Office.Interop.Excel.Range excelCellrange;
            Microsoft.Office.Interop.Excel.Borders border;
            // 
            // Start Excel and get Application object.
            excel = new Microsoft.Office.Interop.Excel.Application();

            // for making Excel visible
            excel.Visible = false;
            excel.DisplayAlerts = false;

            // Creation a new Workbook
            excelworkBook = excel.Workbooks.Add(Type.Missing);

            string root = "////////";
            string extn = "";
            string space = "";
            string temp_path = "";

            int cnt = 0;
            int cnt1 = 0;

            int k = 2;
            int l = 1;
            int m = 2;

            List<String> addpkg = new List<string>();
            List<String> addpkg1 = new List<string>();
            List<String> extnlst = new List<string>();
            List<String> extnlst1 = new List<string>();
            List<String> countval1 = new List<string>();

            string pkg = extract.Split(Path.DirectorySeparatorChar).Last();
            string pkg1 = extract1.Split(Path.DirectorySeparatorChar).Last();

            char[] charsTotrim = { '{', '}', ' ', '=' };

            System.Data.DataTable xaml = new System.Data.DataTable();
            DataRow xamlnull = xaml.NewRow();
            xaml.Columns.Add("XAML");
            xaml.Columns.Add(pkg);
            xaml.Columns.Add(pkg1);
            xaml.Rows.Add(xamlnull);

            System.Data.DataTable fileCnt = new System.Data.DataTable();
            DataRow filenull = fileCnt.NewRow();
            fileCnt.Columns.Add("File_Type");
            fileCnt.Columns.Add("Total_Count");
            fileCnt.Rows.Add(filenull);
            DataRow cntnull = fileCnt.NewRow();
            fileCnt.Rows.Add(cntnull);

            //////////////////////Sheet_For_Tree_View_1 
            excelSheet1 = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.ActiveSheet;
            excelSheet1.Name = "Folder_Structure_1";
            excelSheet1.Cells[1, 1] = pkg;
            foreach (string file in Directory.EnumerateFiles(
            extract, "*.*", SearchOption.AllDirectories))
            {
                if (file.Contains(root))
                    space = new string(' ', root.Length - extract.Length);
                else
                    space = new string(' ', root.Length);
                int position = file.LastIndexOf('\\');
                excelSheet1.Cells[k, l] = file.Replace(root, space).Replace(extract, "");
                root = file.Substring(0, position);
                k++;
                temp_path = file.Substring(position + 1, file.Length - position - 1);
                cnt++;
                int pos = temp_path.LastIndexOf('.');
                extn = temp_path.Substring(pos + 1, temp_path.Length - pos - 1);
                //Console.WriteLine(extn);
                extnlst.Add(extn);
                addpkg.Add(temp_path);
            }
            excelCellrange = excelSheet1.Range[excelSheet1.Cells[1, 1], excelSheet1.Cells[k, l]];
            excelCellrange.EntireColumn.AutoFit();

            var result = extnlst.GroupBy(item => item)
                        .Select(item => new {
                            Name = item.Key,
                            Count = item.Count()
                        })
            .OrderByDescending(item => item.Count)
            .ThenBy(item => item.Name);

            DataRow titlerow = fileCnt.NewRow();
            titlerow["File_Type"] = pkg.ToUpper();
            fileCnt.Rows.Add(titlerow);
            foreach (var addex in result)
            {
                DataRow dttcount = fileCnt.NewRow();
                string[] splitVal = addex.ToString().Split(',');
                dttcount["File_Type"] = "." + splitVal[0].Replace("Name", "").Trim(charsTotrim);
                dttcount["Total_Count"] = splitVal[1].Replace("Count", "").Trim(charsTotrim);
                fileCnt.Rows.Add(dttcount);
            }
            DataRow cntnull1 = fileCnt.NewRow();
            fileCnt.Rows.Add(cntnull1);

            ////////////////////////////Sheet For Tree_View_2
            excelSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.Sheets.Add(System.Reflection.Missing.Value,
                 excelworkBook.Worksheets[excelworkBook.Worksheets.Count],
                 System.Reflection.Missing.Value,
                 System.Reflection.Missing.Value);
            excelSheet2.Name = "Folder_Structure_2";
            excelSheet2.Cells[1, 1] = pkg1;
            foreach (string file in Directory.EnumerateFiles(
           extract1, "*.*", SearchOption.AllDirectories))
            {
                if (file.Contains(root))
                    space = new string(' ', root.Length - extract1.Length);
                else
                    space = new string(' ', root.Length);
                int position = file.LastIndexOf('\\');
                excelSheet2.Cells[m, l] = file.Replace(root, space).Replace(extract1, "");
                root = file.Substring(0, position);
                m++;
                temp_path = file.Substring(position + 1, file.Length - position - 1);
                cnt1++;
                int pos = temp_path.LastIndexOf('.');
                extn = temp_path.Substring(pos + 1, temp_path.Length - pos - 1);
                extnlst1.Add(extn);
                addpkg1.Add(temp_path);
            }
            excelCellrange = excelSheet2.Range[excelSheet2.Cells[1, 1], excelSheet2.Cells[k, l]];
            excelCellrange.EntireColumn.AutoFit();

            var result1 = extnlst1.GroupBy(item => item)
                        .Select(item => new {
                            Name = item.Key,
                            Count = item.Count()
                        })
            .OrderByDescending(item => item.Count)
            .ThenBy(item => item.Name);

            DataRow titlerow1 = fileCnt.NewRow();
            titlerow1["File_Type"] = pkg1.ToUpper();
            fileCnt.Rows.Add(titlerow1);
            foreach (var addex in result1)
            {
                DataRow dttcount = fileCnt.NewRow();
                string[] splitVal = addex.ToString().Split(',');
                dttcount["File_Type"] = "." + splitVal[0].Replace("Name", "").Trim(charsTotrim);
                dttcount["Total_Count"] = splitVal[1].Replace("Count", "").Trim(charsTotrim);
                fileCnt.Rows.Add(dttcount);
            }

            var finalpkg = addpkg.Union(addpkg1);
            foreach (var check in finalpkg)
            {
                DataRow dtxaml = xaml.NewRow();
                dtxaml["XAML"] = check;
                if (addpkg.Contains(check))
                {
                    dtxaml[pkg] = "Yes";
                }
                else
                {
                    dtxaml[pkg] = "No";
                }

                if (addpkg1.Contains(check))
                {
                    dtxaml[pkg1] = "Yes";
                }
                else
                {
                    dtxaml[pkg1] = "No";
                }
                xaml.Rows.Add(dtxaml);
            }
            DataRow dtcountnull = xaml.NewRow();
            xaml.Rows.Add(dtcountnull);
            DataRow dtcount = xaml.NewRow();
            dtcount["XAML"] = "Total Number of Files";
            dtcount[pkg] = cnt;
            dtcount[pkg1] = cnt1;
            xaml.Rows.Add(dtcount);

            /////////////////////////////Sheet For List of File Counts in a Folder
            excelSheet3 = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.Sheets.Add(System.Reflection.Missing.Value,
                 excelworkBook.Worksheets[excelworkBook.Worksheets.Count],
                 System.Reflection.Missing.Value,
                 System.Reflection.Missing.Value);
            excelSheet3.Name = "File_Count";
            int p = 1;
            int q = 1;
            foreach (DataRow dr in fileCnt.Rows)
            {
                foreach (DataColumn dc in fileCnt.Columns)
                {
                    if (p == 1)
                        excelSheet3.Cells[p, q] = dc.ColumnName.ToString();
                    else
                        excelSheet3.Cells[p, q] = dr[dc].ToString();
                    q++;
                }
                p++;
                q = 1;
            }
            //put everything in a table
            if (fileCnt.Rows.Count == 0)
                excelCellrange = excelSheet3.Range[excelSheet3.Cells[1, 1], excelSheet3.Cells[1, fileCnt.Columns.Count]];
            else
                excelCellrange = excelSheet3.Range[excelSheet3.Cells[1, 1], excelSheet3.Cells[fileCnt.Rows.Count, fileCnt.Columns.Count]];
            excelCellrange.EntireColumn.AutoFit();
            excelCellrange.EntireRow.AutoFit();
            border = excelCellrange.Borders;
            border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            border.Weight = 2d;
            excelCellrange = excelSheet3.Range["A1", "B1"];
            excelCellrange.Interior.Color = System.Drawing.
            ColorTranslator.ToOle(System.Drawing.Color.Lavender);
            excelCellrange.Font.Color = System.Drawing.
            ColorTranslator.ToOle(System.Drawing.Color.Black);
            excelCellrange.Font.Bold = true;

            /////////////////////////////Sheet For Compare Sheet
            excelSheet4 = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.Sheets.Add(System.Reflection.Missing.Value,
                 excelworkBook.Worksheets[excelworkBook.Worksheets.Count],
                 System.Reflection.Missing.Value,
                 System.Reflection.Missing.Value);
            excelSheet4.Name = "Compare_Sheet";
            //excelCellrange = new Microsoft.Office.Interop.Excel.Range();
            int i = 1;
            int j = 1;
            foreach (DataRow dr in xaml.Rows)
            {
                foreach (DataColumn dc in xaml.Columns)
                {
                    if (i == 1)
                        excelSheet4.Cells[i, j] = dc.ColumnName.ToString();
                    else
                        excelSheet4.Cells[i, j] = dr[dc].ToString();
                    j++;
                }
                i++;
                j = 1;
            }
            //put everything in a table
            if (xaml.Rows.Count == 0)
                excelCellrange = excelSheet4.Range[excelSheet4.Cells[1, 1], excelSheet4.Cells[1, xaml.Columns.Count]];
            else
                excelCellrange = excelSheet4.Range[excelSheet4.Cells[1, 1], excelSheet4.Cells[xaml.Rows.Count, xaml.Columns.Count]];
            excelCellrange.EntireColumn.AutoFit();
            excelCellrange.EntireRow.AutoFit();
            border = excelCellrange.Borders;
            border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            border.Weight = 2d;
            excelCellrange = excelSheet4.Range["A1", "C1"];
            excelCellrange.Interior.Color = System.Drawing.
            ColorTranslator.ToOle(System.Drawing.Color.Lavender);
            excelCellrange.Font.Color = System.Drawing.
            ColorTranslator.ToOle(System.Drawing.Color.Black);
            excelCellrange.Font.Bold = true;

            excelworkBook.SaveAs(outputPath);
            string rslt = "File Created";
            Output.Set(context, rslt);
        }
    }
}
