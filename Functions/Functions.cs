using ClosedXML.Excel;
using ExcelInterface.Attributes;
using ExcelInterface.Models;
using System.Globalization;
using System.Reflection;
using System.Text;

namespace ExcelBridge
{
    public class Functions
    {
        public static void FromObjectOfType<T>(T Origin, string TemplateName, string TargetName) where T : new()
        {
            var templatecontent = File.ReadAllBytes(TemplateName);
            var retbytes = FromObjectOfType<T>(Origin, templatecontent);
            File.WriteAllBytes(TargetName, retbytes);
        }

        public static byte[] FromObjectOfType<T>(T Origin, byte[] TemplateContent) where T : new()
        {
            byte[] ret = null;
            using (var workbook = new XLWorkbook(new MemoryStream(TemplateContent)))
            {
                Type ttype = typeof(T);
                var tprops = ttype.GetProperties();
                foreach (var property in tprops)
                {

                    foreach (var attribute in property.GetCustomAttributes(false))
                    {
                        if (attribute.GetType() == typeof(EIBase))
                        {

                            EIBase ebo = (attribute as EIBase);
                        }

                        if (attribute.GetType() == typeof(EIRegister))
                        {
                            EIRegister eco = (attribute as EIRegister);

                            var errorlist = new StringBuilder();
                            var infolist = new StringBuilder();
                            Type itemType = property.PropertyType;

                            MethodInfo writeecl = typeof(Functions).GetMethod("ECOWrite");
                            MethodInfo genericwrite = writeecl.MakeGenericMethod(itemType);

                            var originobject = property.GetValue(Origin);
                            if (originobject != null)
                                genericwrite.Invoke(null, new object[] { workbook, attribute, originobject });
                        }

                        if (attribute.GetType() == typeof(EIList))
                        {
                            var errorlist = new StringBuilder();
                            var infolist = new StringBuilder();
                            Type itemType = null;
                            Type type = property.PropertyType;

                            if (type.IsGenericType && type.GetGenericTypeDefinition()
                                    == typeof(List<>))
                            {
                                itemType = type.GetGenericArguments()[0]; // use this...
                            }

                            if (itemType == null)
                                throw new Exception("EIList must be a List");

                            MethodInfo writeecl = typeof(Functions).GetMethod("WriteECL");
                            MethodInfo genericwrite = writeecl.MakeGenericMethod(itemType);

                            var originobject = property.GetValue(Origin);
                            if (originobject != null)
                                genericwrite.Invoke(null, new object[] { workbook, attribute, originobject });
                        }
                    }
                }
                var m = new MemoryStream();
                workbook.SaveAs(m);
                ret = m.ToArray();
            }

            return ret;
        }


        public static T ToObjectOfType<T>(string FileName) where T : new()
        {
            var b = File.ReadAllBytes(FileName);
            return ToObjectOfType<T>(File.ReadAllBytes(FileName));
            T ret = new T();
            using (var workbook = new XLWorkbook(FileName, new LoadOptions {/* ReadOnly = false*/ }))
            {
            }
            using (var stream = File.OpenRead(FileName))
            {
                ret = ToObjectOfType<T>(stream);
            }
            return ret;
        }
        public static T ToObjectOfType<T>(byte[] FileContent) where T : new()
        {
            T ret = new T();
            //    using (var workbook = new XLWorkbook(new MemoryStream(FileContent), new LoadOptions { ReadOnly = true }))
            using (var stream = new MemoryStream(FileContent))
            {
                ret = ToObjectOfType<T>(stream);
            }
            return ret;
        }
        public static T ToObjectOfType<T>(Stream Content) where T : new()
        {
            T ret = new T();
            using (var workbook = new XLWorkbook(Content, new LoadOptions { /* ReadOnly = false */}))
            {
                ret = ToObjectOfType<T>(workbook);
            }
            return ret;
        }
        private static T ToObjectOfType<T>(XLWorkbook workbook) where T : new()
        {
            T ret = new T();
            Type ttype = typeof(T);
            var tprops = ttype.GetProperties();
            foreach (var property in tprops)
            {

                foreach (var attribute in property.GetCustomAttributes(false))
                {
                    if (attribute.GetType() == typeof(EIBase))
                    {

                        EIBase ebo = (attribute as EIBase);
                        IXLWorksheet worksheet = null;
                        if (ebo.SheetName != null)
                            worksheet = workbook.Worksheet(ebo.SheetName);
                        if (ebo.SheetPosition > 0)
                            worksheet = workbook.Worksheet(ebo.SheetPosition);
                        var cell = worksheet.Cell(ebo.Position);

                        ConvertCellToObject(cell, property, ret);

                    }



                    if (attribute.GetType() == typeof(EIRegister))
                    {
                        var targetobject = property.GetValue(ret);
                        if (targetobject == null)
                        {
                            var type = property.PropertyType;
                            targetobject = Activator.CreateInstance(type);
                            property.SetValue(ret, targetobject);
                        }
                        EIRegister eco = (attribute as EIRegister);
                        IXLWorksheet worksheet = null;
                        if (eco.SheetName != null)
                            worksheet = workbook.Worksheet(eco.SheetName);
                        if (eco.SheetPosition > 0)
                            worksheet = workbook.Worksheet(eco.SheetPosition);
                        var startcell = worksheet.Cell(eco.StartPosition);
                        var endcell = worksheet.Cell(eco.EndPosition);
                        for (var row = startcell.Address.RowNumber; row <= endcell.Address.RowNumber; row++)
                        {
                            //Let's keep it simple for now
                            var headercol = startcell.Address.ColumnNumber;
                            var datacol = endcell.Address.ColumnNumber;


                            var subtype = property.PropertyType;
                            var subtprops = subtype.GetProperties();
                            foreach (var subproperty in subtprops)
                            {
                                var ecomemberprop = subproperty.GetCustomAttributes(false).Where(p => p.GetType() == typeof(EIRegisterMember)).FirstOrDefault() as EIRegisterMember;
                                if (ecomemberprop != null)
                                {
                                    //var subobject= subproperty.GetValue()
                                    var targetheader = subproperty.Name;
                                    if (!string.IsNullOrEmpty(ecomemberprop.Header))
                                        targetheader = ecomemberprop.Header;
                                    var headercellvalue = worksheet.Cell(row, headercol).GetString();
                                    var datacell = worksheet.Cell(row, datacol);
                                    if (headercellvalue.Contains(targetheader)) //Keep it simple here as well, has to change though.
                                    {
                                        ConvertCellToObject(datacell, subproperty, targetobject);


                                    }
                                }

                            }
                        }


                    }


                    if (attribute.GetType() == typeof(EIList))
                    {
                        var targetobject = property.GetValue(ret);
                        if (targetobject == null)
                        {
                            var type = property.PropertyType;
                            targetobject = Activator.CreateInstance(type);
                            property.SetValue(ret, targetobject);
                        }
                        EIList ecl = (attribute as EIList);
                        IXLWorksheet worksheet = null;
                        if (ecl.SheetName != null && workbook.Worksheets.Contains(ecl.SheetName))
                            worksheet = workbook.Worksheet(ecl.SheetName);
                        if (ecl.SheetPosition > 0)
                            worksheet = workbook.Worksheet(ecl.SheetPosition);
                        if (worksheet == null)
                            continue;
                        ReadECL(worksheet, property, targetobject);
                        //SheetAttribute sa = (a as SheetAttribute);
                        //SheetName = sa.Name;
                        //SheetPosition = sa.Position;
                        //HeaderRow = sa.HeaderRow;
                    }
                }
            }

            return ret;
        }



        public static SpreadSheetMeta GetSpreadsheetMeta(EIList eclinfo, Type t)
        {
            var ret = new SpreadSheetMeta();
            ret.SheetName = eclinfo.SheetName;
            ret.SheetPosition = eclinfo.SheetPosition;
            ret.HeaderRow = eclinfo.HeaderRow;

            foreach (PropertyInfo p in t.GetProperties())
            {
                var c = new ColumnMeta();
                c.pi = p;
                c.type = p.PropertyType;
                // for every property loop through all attributes
                foreach (Attribute a in p.GetCustomAttributes(false))
                {
                    if (a.GetType() == typeof(EIRegisterMember))
                    {
                        EIRegisterMember sa = (a as EIRegisterMember);
                        c.ColumnName = sa.Header;
                    }
                }
                ret.Columns.Add(c);
            }
            return ret;

        }

        public static string ReplaceInArray(string origin, byte[] bytesToSearch, byte replacement)
        {
            if (origin == null)
                return null;
            if (origin == "")
                return "";
            var ret = new byte[origin.Length];
            var retspan = new Span<byte>(ret);
            var originspan = UTF8Encoding.UTF8.GetBytes(origin).AsSpan<byte>();
            var bytestosearchspan = bytesToSearch.AsSpan<byte>();
            var targetpos = 0;
            for (int i = 0; i < origin.Length; i++)
            {
                if (origin.Length - i >= bytesToSearch.Length && bytestosearchspan.SequenceEqual(originspan.Slice(i, bytesToSearch.Length)))
                {
                    retspan[targetpos++] = replacement;
                    i += bytesToSearch.Length - 1;
                }
                else
                    retspan[targetpos++] = originspan[i];

            }
            return UTF8Encoding.UTF8.GetString(retspan.Slice(0, targetpos));//
        }
        public static void WriteECL<T>(XLWorkbook wb, EIList eclinfo, List<T> Input) where T : new()
        {
            var meta = GetSpreadsheetMeta(eclinfo, typeof(T));
            //using (XLWorkbook wb = new XLWorkbook())
            {
                IXLWorksheet ws = null;
                if (meta.SheetName != null)
                {
                    if (wb.Worksheets.Contains(meta.SheetName))
                        ws = wb.Worksheet(meta.SheetName);
                    else
                        ws = wb.Worksheets.Add(meta.SheetName);
                }
                if (meta.SheetPosition > 0)
                {
                    if (wb.Worksheets.Count >= meta.SheetPosition)
                        ws = wb.Worksheet(meta.SheetPosition);
                    else
                        ws = wb.Worksheets.Add($"New Sheet Pos {meta.SheetPosition}");
                }
                if (ws == null)
                    return;
                var row = meta.HeaderRow;
                var colnum = eclinfo.StartColumn;
                foreach (var col in meta.Columns)
                {
                    col.ColumnNumber = ++colnum;
                    ws.Cell(row, col.ColumnNumber).SetValue<string>(col.ColumnName);
                }
                row++;
                foreach (var inputrow in Input)
                {
                    foreach (var col in meta.Columns)
                    {
                        var value = col.pi.GetValue(inputrow);


                        switch (col.type.ToString())
                        {
                            case "System.String":
                                {
                                    //var svalue = (string)value;
                                    //if (svalue.Length > 0 && int.TryParse(svalue, out int n))
                                    //    value = $"'{value}";
                                    var sanitizedValue = ReplaceInArray((string)value, new byte[] { 0xef, 0xbf, 0xbd }, 0x27);
                                    ws.Cell(row, col.ColumnNumber).SetValue<string>(sanitizedValue);
                                    break;
                                }
                            case "System.Decimal":
                            case "System.Nullable`1[System.Decimal]":
                                {
                                    var doublevalue = 0.0;
                                    try
                                    {
                                        doublevalue = Convert.ToDouble(value);

                                    }
                                    catch { };
                                    ws.Cell(row, col.ColumnNumber).SetValue<double>(doublevalue);


                                    break;
                                }
                            case "System.Int32":
                            case "System.Nullable`1[System.Int32]":
                                {
                                    var intvalue = 0;
                                    try
                                    {
                                        intvalue = Convert.ToInt32(value);

                                    }
                                    catch { };
                                    ws.Cell(row, col.ColumnNumber).SetValue<int>(intvalue);


                                    break;
                                }
                            case "System.DateTime":
                            case "System.Nullable`1[System.DateTime]":
                                {
                                    if ((DateTime)value != DateTime.MinValue)
                                        ws.Cell(row, col.ColumnNumber).SetValue<DateTime>((DateTime)value);
                                    //DateTime? value = null;
                                    //try
                                    //{
                                    //    if (cell.Value.GetType() == typeof(string))
                                    //    {
                                    //        DateTime temp;
                                    //        DateTime.TryParseExact(cell.GetString(),
                                    //                               "yyyy/dd/MM",
                                    //                               CultureInfo.InvariantCulture,
                                    //                               DateTimeStyles.None,
                                    //                               out temp);
                                    //        value = temp;
                                    //        //DateTime.TryParse(cell.GetString(), temp);
                                    //    }
                                    //    else if (cell.Value.GetType() == typeof(double))
                                    //    {
                                    //        double d = cell.GetDouble();
                                    //        value = DateTime.FromOADate(d);
                                    //    }
                                    //    else
                                    //        value = cell.GetDateTime();
                                    //    colinfo.pi.SetValue(res, value, null);

                                    //}
                                    //catch { };
                                    //colinfo.pi.SetValue(res, value, null);
                                    break;
                                }
                            default:
                                {
                                    //Console.WriteLine($"{colinfo.type.ToString()} type not found");
                                    break;
                                }
                        }



                    }
                    row++;

                }



                //wb.SaveAs(FileName);

            }


        }

        public static void ReadECL(IXLWorksheet worksheet, PropertyInfo pi, object target)
        {
            string SheetName = "";
            int SheetPosition = 0;
            int HeaderRow = 0;

            var errorlist = new StringBuilder();
            var infolist = new StringBuilder();
            Type itemType = null;
            Type type = pi.PropertyType;
            if (type.IsGenericType && type.GetGenericTypeDefinition()
                    == typeof(List<>))
            {
                itemType = type.GetGenericArguments()[0]; // use this...
            }
            if (itemType == null)
                throw new Exception("EIList must be a List");



            foreach (Attribute a in pi.GetCustomAttributes(false))
            {

                if (a.GetType() == typeof(EIList))
                {
                    EIList sa = (a as EIList);
                    SheetName = sa.SheetName;
                    SheetPosition = sa.SheetPosition;
                    HeaderRow = sa.HeaderRow;
                }

            }

            var columnList = new List<EIListColumn>();
            bool hasmustbepopulatedcolumn = false;
            foreach (PropertyInfo p in itemType.GetProperties())
            {
                var c = new EIListColumn();
                c.pi = p;
                c.type = p.PropertyType;
                // for every property loop through all attributes
                foreach (Attribute a in p.GetCustomAttributes(false))
                {
                    if (a.GetType() == typeof(EIRegisterMember))
                    {
                        EIRegisterMember sa = (a as EIRegisterMember);
                        c.ColumnName = sa.Header;
                        c.DateTimeFormat = sa.DateTimeFormat;
                        c.MustBePopulated = sa.MustBePopulated;
                        c.Trim = sa.Trim;
                        if (c.MustBePopulated)
                            hasmustbepopulatedcolumn = true;
                        c.CanBeOmitted = sa.CanBeOmitted;
                    }
                }
                columnList.Add(c);
            }

            var lastrow = worksheet.LastCellUsed().Address.RowNumber;
            var lastcol = worksheet.LastCellUsed().Address.ColumnNumber;
            for (var c1 = 1; c1 <= lastcol; c1++)
            {
                string targetName = worksheet.Cell(HeaderRow, c1).GetString();
                var col = columnList.Where(c => c.ColumnName == targetName).FirstOrDefault();
                if (col != null)
                {
                    col.ColumnNumber = c1;
                    col.Present = true;
                }
                else
                    infolist.AppendFormat($"Column {targetName} not found\n");
                //Console.WriteLine($"Column {targetName} not found");
            }
            var notfoundcolumns = columnList.Where(c => c.ColumnNumber == 0 && c.CanBeOmitted == false).Select(c => c.ColumnName).ToList();
            if (notfoundcolumns.Count > 0)
                errorlist.AppendFormat($"Column(s) {notfoundcolumns.Aggregate((x, y) => x + "," + y)} not found\n");

            if (errorlist.Length > 0)
                throw new Exception(errorlist.ToString());

            for (var row = HeaderRow + 1; row <= lastrow; row++)
            {
                bool empty = true;
                var res = Activator.CreateInstance(itemType);
                //T res = new T();
                for (var col = 1; col <= lastcol; col++)
                {
                    var colinfo = columnList.Where(c => c.ColumnNumber == col).FirstOrDefault();
                    if (colinfo != null)
                    {
                        if (!colinfo.Present)
                            continue;
                        var cell = worksheet.Cell(row, col);
                        switch (colinfo.type.ToString())
                        {
                            case "System.String":
                                {

                                    var value = "";
                                    if (string.IsNullOrEmpty(cell.FormulaA1))
                                        value = cell.GetString();
                                    else
                                        value = (string)cell.CachedValue;
                                    if (colinfo.MustBePopulated && !string.IsNullOrEmpty(value))
                                        empty = false;
                                    if (colinfo.Trim)
                                        value = value.Trim();
                                    var sanitizedValue = ReplaceInArray((string)value, new byte[] { 0xef, 0xbf, 0xbd }, 0x27);

                                    colinfo.pi.SetValue(res, sanitizedValue, null);
                                    break;
                                }
                            case "System.Decimal":
                            case "System.Nullable`1[System.Decimal]":
                                {

                                    var value = 0M;
                                    try
                                    {
                                        if (string.IsNullOrEmpty(cell.FormulaA1))
                                            value = Convert.ToDecimal(cell.GetDouble());
                                        else
                                        {
                                            if (cell.CachedValue.GetType() == typeof(string))
                                            {
                                                double dvalue;

                                                double.TryParse((string)cell.CachedValue, out dvalue);
                                                value = Convert.ToDecimal(dvalue);

                                            }
                                            else
                                                value = Convert.ToDecimal(cell.CachedValue);
                                        }


                                        //if (colinfo.MustBePopulated && value!=null)
                                        //    empty = false;

                                    }
                                    catch
                                    {

                                    };
                                    colinfo.pi.SetValue(res, value, null);
                                    break;
                                }
                            case "System.DateTime":
                            case "System.Nullable`1[System.DateTime]":
                                {
                                    DateTime? value = null;
                                    try
                                    {
                                        if (cell.Value.GetType() == typeof(string))
                                        {
                                            var dateformat = "yyyy/dd/MM";
                                            if (!string.IsNullOrEmpty(colinfo.DateTimeFormat))
                                                dateformat = colinfo.DateTimeFormat;
                                            DateTime temp;
                                            DateTime.TryParseExact(cell.GetString(),
                                                                   dateformat,
                                                                   CultureInfo.InvariantCulture,
                                                                   DateTimeStyles.None,
                                                                   out temp);
                                            value = temp;
                                            //DateTime.TryParse(cell.GetString(), temp);
                                        }
                                        else if (cell.Value.GetType() == typeof(double))
                                        {
                                            double d = cell.GetDouble();
                                            if (!string.IsNullOrEmpty(colinfo.DateTimeFormat))
                                            {
                                                var dateformat = colinfo.DateTimeFormat;
                                                DateTime temp;
                                                DateTime.TryParseExact(cell.GetString(),
                                                                       dateformat,
                                                                       CultureInfo.InvariantCulture,
                                                                       DateTimeStyles.None,
                                                                       out temp);
                                                value = temp;

                                            }
                                            else
                                                value = DateTime.FromOADate(d);
                                        }
                                        else
                                            value = cell.GetDateTime();
                                        colinfo.pi.SetValue(res, value, null);

                                    }
                                    catch { };
                                    colinfo.pi.SetValue(res, value, null);
                                    break;
                                }
                            default:
                                {
                                    Console.WriteLine($"{colinfo.type.ToString()} type not found");
                                    break;
                                }
                        }

                    }

                }
                if (hasmustbepopulatedcolumn)
                {
                    if (!empty)
                    {
                        target.GetType().GetMethod("Add").Invoke(target, new[] { res });
                    }
                }
                else
                    target.GetType().GetMethod("Add").Invoke(target, new[] { res });
            }
        }

        public static void ECOWrite<T>(XLWorkbook wb, EIRegister ecoinfo, T Input) where T : new()
        {
            //var meta = GetSpreadsheetMeta(eclinfo, typeof(T));
            //using (XLWorkbook wb = new XLWorkbook())
            {
                IXLWorksheet ws = null;
                if (ecoinfo.SheetName != null)
                {
                    if (wb.Worksheets.Contains(ecoinfo.SheetName))
                        ws = wb.Worksheet(ecoinfo.SheetName);
                    else
                        ws = wb.Worksheets.Add(ecoinfo.SheetName);
                }
                if (ecoinfo.SheetPosition > 0)
                {
                    if (wb.Worksheets.Count >= ecoinfo.SheetPosition)
                        ws = wb.Worksheet(ecoinfo.SheetPosition);
                    else
                        ws = wb.Worksheets.Add($"New Sheet Pos {ecoinfo.SheetPosition}");
                }
                if (ws == null)
                    return;


                var startcell = ws.Cell(ecoinfo.StartPosition);
                var endcell = ws.Cell(ecoinfo.EndPosition);

                var headercol = startcell.Address.ColumnNumber;
                var datacol = endcell.Address.ColumnNumber;
                var row = startcell.Address.RowNumber;
                foreach (PropertyInfo pi in typeof(T).GetProperties())
                {
                    EIRegisterMember sa = (EIRegisterMember)pi.GetCustomAttributes(false).Where(p => p.GetType() == typeof(EIRegisterMember)).FirstOrDefault();
                    //foreach (Attribute a in p.GetCustomAttributes(false))
                    //{
                    //    if (a.GetType() == typeof(ECOMember))
                    //    {
                    //        ECOMember sa = (a as ECOMember);
                    //        c.ColumnName = sa.Header;
                    //    }
                    //}
                    if (sa == null)
                        continue;


                    ws.Cell(row, headercol).SetValue<string>((string)sa.Header);
                    var value = pi.GetValue(Input);


                    switch (value.GetType().ToString())
                    {
                        case "System.String":
                            {
                                //var svalue = (string)value;
                                //if (svalue.Length > 0 && int.TryParse(svalue, out int n))
                                //    value = $"'{value}";

                                var sanitizedValue = ReplaceInArray((string)value, new byte[] { 0xef, 0xbf, 0xbd }, 0x27);

                                ws.Cell(row, datacol).SetValue<string>(sanitizedValue);
                                break;
                            }
                        case "System.Decimal":
                        case "System.Nullable`1[System.Decimal]":
                            {
                                var doublevalue = 0.0;
                                try
                                {
                                    doublevalue = Convert.ToDouble(value);

                                }
                                catch { };
                                ws.Cell(row, datacol).SetValue<double>(doublevalue);


                                break;
                            }
                        case "System.Int32":
                        case "System.Nullable`1[System.Int32]":
                            {
                                var intvalue = 0;
                                try
                                {
                                    intvalue = Convert.ToInt32(value);

                                }
                                catch { };
                                ws.Cell(row, datacol).SetValue<int>(intvalue);


                                break;
                            }
                        case "System.DateTime":
                        case "System.Nullable`1[System.DateTime]":
                            {
                                if ((DateTime)value != DateTime.MinValue)
                                    ws.Cell(row, datacol).SetValue<DateTime>((DateTime)value);
                                //DateTime? value = null;
                                //try
                                //{
                                //    if (cell.Value.GetType() == typeof(string))
                                //    {
                                //        DateTime temp;
                                //        DateTime.TryParseExact(cell.GetString(),
                                //                               "yyyy/dd/MM",
                                //                               CultureInfo.InvariantCulture,
                                //                               DateTimeStyles.None,
                                //                               out temp);
                                //        value = temp;
                                //        //DateTime.TryParse(cell.GetString(), temp);
                                //    }
                                //    else if (cell.Value.GetType() == typeof(double))
                                //    {
                                //        double d = cell.GetDouble();
                                //        value = DateTime.FromOADate(d);
                                //    }
                                //    else
                                //        value = cell.GetDateTime();
                                //    colinfo.pi.SetValue(res, value, null);

                                //}
                                //catch { };
                                //colinfo.pi.SetValue(res, value, null);
                                break;
                            }
                        default:
                            {
                                //Console.WriteLine($"{colinfo.type.ToString()} type not found");
                                break;
                            }
                    }


                    row++;
                }


                //wb.SaveAs(FileName);

            }


        }


        public static void ConvertCellToObject(IXLCell cell, PropertyInfo pi, object target)
        {
            var ecomemberprop = pi.GetCustomAttributes(false).Where(p => p.GetType() == typeof(EIRegisterMember)).FirstOrDefault() as EIRegisterMember;
            var eboprop = pi.GetCustomAttributes(false).Where(p => p.GetType() == typeof(EIBase)).FirstOrDefault() as EIBase;
            switch (pi.PropertyType.ToString())
            {
                case "System.String":
                    {
                        var value = cell.GetString();
                        if (ecomemberprop != null)
                        {
                            if (ecomemberprop.Trim)
                                value = value.Trim();

                        }
                        //if (colinfo.MustBePopulated && !string.IsNullOrEmpty(value))
                        //    empty = false;
                        var sanitizedValue = ReplaceInArray((string)value, new byte[] { 0xef, 0xbf, 0xbd }, 0x27);

                        pi.SetValue(target, sanitizedValue, null);
                        break;
                    }
                case "System.Decimal":
                case "System.Nullable`1[System.Decimal]":
                    {
                        var value = 0M;
                        try
                        {
                            value = Convert.ToDecimal(cell.GetDouble());
                            //if (colinfo.MustBePopulated && value!=null)
                            //    empty = false;

                        }
                        catch
                        {

                        };
                        pi.SetValue(target, value, null);
                        break;
                    }
                case "System.DateTime":
                case "System.Nullable`1[System.DateTime]":
                    {
                        DateTime? value = null;
                        try
                        {
                            if (cell.Value.GetType() == typeof(string))
                            {
                                var dateformat = "yyyy/dd/MM";

                                if (ecomemberprop != null && !string.IsNullOrEmpty(ecomemberprop.DateTimeFormat))
                                    dateformat = ecomemberprop.DateTimeFormat;
                                if (eboprop != null && !string.IsNullOrEmpty(eboprop.DateTimeFormat))
                                    dateformat = eboprop.DateTimeFormat;
                                DateTime temp;
                                DateTime.TryParseExact(cell.GetString(),
                                                       dateformat,
                                                       CultureInfo.InvariantCulture,
                                                       DateTimeStyles.None,
                                                       out temp);
                                value = temp;
                                //DateTime.TryParse(cell.GetString(), temp);
                            }
                            else if (cell.Value.GetType() == typeof(double))
                            {
                                double d = cell.GetDouble();
                                value = DateTime.FromOADate(d);
                            }
                            else
                                value = cell.GetDateTime();
                            pi.SetValue(target, value, null);

                        }
                        catch { };
                        //pi.SetValue(target, value, null);
                        break;
                    }
                default:
                    {
                        Console.WriteLine($"{pi.GetType().ToString()} type not found");
                        break;
                    }
            }
        }
    }
}