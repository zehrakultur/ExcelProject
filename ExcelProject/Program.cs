using OfficeOpenXml;

namespace ExcelProject
{

    class Program
    {
        static void Main(string[] args)
        {
            string file = @"C:\Users\Acer\Desktop\TEST-01\EKSTRE-GIRDI.xlsx";
            string formulFile = @"C:\Users\Acer\Desktop\TEST-01\FORMUL-GIRDI.xlsx";
            

            using(ExcelPackage package = new ExcelPackage(new FileInfo(file)))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                var sheet = package.Workbook.Worksheets["Sayfa1"];
                var girdi = new Program().GetList<Girdi>(sheet);
            }
            using (ExcelPackage package = new ExcelPackage(new FileInfo(formulFile)))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                var sheetFormul = package.Workbook.Worksheets["FORMUL"];
                var formuller = new Program().FormulList<Formul>(sheetFormul);
            }
        }

        private List<T> GetList<T>(ExcelWorksheet sheet)
        {
            List<T> list= new List<T>();
            var columnInfo = Enumerable.Range(1, sheet.Dimension.Columns).ToList().Select(n =>

            new {Index=n, ColumnName=sheet.Cells[1,n].Value.ToString()}
            );

            for(int row=2; row<sheet.Dimension.Rows; row++)
            {
                T obj = (T)Activator.CreateInstance(typeof(T));
                foreach(var prop in typeof(T).GetProperties())
                {
                    int col = columnInfo.SingleOrDefault(c => c.ColumnName == prop.Name).Index;
                    var val = sheet.Cells[row, col].Value;
                    var propType = prop.PropertyType;
                    prop.SetValue(obj, Convert.ChangeType(val, propType));
                }
                list.Add(obj);
            }

            return list;
        }

        private List<T> FormulList<T>(ExcelWorksheet sheet)
        {
            List<T> listFormul = new List<T>();
            var columnInfo = Enumerable.Range(1, sheet.Dimension.Columns).ToList().Select(n =>

            new { Index = n, ColumnName = sheet.Cells[1, n].Value.ToString() }
            );

            for (int row = 2; row < sheet.Dimension.Rows; row++)
            {
                T obj = (T)Activator.CreateInstance(typeof(T));
                foreach (var prop in typeof(T).GetProperties())
                {
                    int col = columnInfo.SingleOrDefault(c => c.ColumnName == prop.Name).Index;
                    var val = sheet.Cells[row, col].Value;
                    var propType = prop.PropertyType;
                    prop.SetValue(obj, Convert.ChangeType(val, propType));
                }
                listFormul.Add(obj);
            }

            return listFormul;
        }
    }
        public class Girdi
        {
            public string SIRA_NO { get; set; }
            public string ADET { get; set;}
            public string KG_DESI { get; set; }
            public string MESAFE { get; set; }

        }

        public class Formul
        {
            public string DESİ { get; set; }
            public string KISA_ŞEHİRİÇİ_YAKIN { get; set; }
            public string UZAK_ORTA { get; set; }
        }


}
