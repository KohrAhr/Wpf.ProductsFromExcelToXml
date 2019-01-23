using ExcelToXML.Types;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelToXML.Functions;
using ExcelObject = Microsoft.Office.Interop.Excel;

namespace ExcelToXML.Core
{
    public static class CoreFunctions
    {
        public static string GatherProductInformation(ProductBlock product, ExcelFile excelFile)
        {
            string result = "";

            // now cycle. find all and include all one by one
            int row = product.start;

            ExcelObject.Worksheet worksheet = excelFile.sheet.Worksheets[product.worksheet];
            do
            {
                if (IsHeaderInThisRow(worksheet, row))
                {
                    // Start!
                    do
                    {
                        row++;

                        string id = ExcelFunctions.GetString(worksheet, row, 1);

                        // Is number in first col? Yes, mean new option
                        if (!String.IsNullOrEmpty(id))
                        {
                            int counter = 0;

                            result += "\n\t<product>";

                            result += "\n\t\t<category>" + product.worksheetName + "</category>";
                            result += "\n\t\t<subcategory>" + product.name + "</subcategory>";
                            result += "\n\t\t<name>" + ExcelFunctions.GetString(worksheet, row, 2) + "</name>";
                            string description = ExcelFunctions.GetString(worksheet, row + 1, 2);
                            if (!String.IsNullOrEmpty(description))
                            {
                                result += "\n\t\t<description>" + description.Trim() + "</description>";
                            }

                            // now while match pattern collect data
                            result += "\n\t\t<suboptions>";

                            do
                            {
                                counter++;

                                string o1 = ExcelFunctions.GetString(worksheet, row, 1);
                                string o2 = ExcelFunctions.GetString(worksheet, row, 3);
                                string o3 = ExcelFunctions.GetString(worksheet, row, 4);
                                // price for ptBed. weight for ptRegular
                                string o4 = ExcelFunctions.GetString(worksheet, row, 5);
                                // price for ptRegular
                                string o5 = ExcelFunctions.GetString(worksheet, row, 6);

                                if ((!String.IsNullOrEmpty(o1) && counter > 1) || String.IsNullOrEmpty(o2) || String.IsNullOrEmpty(o3) || String.IsNullOrEmpty(o4) || (String.IsNullOrEmpty(o5) && product.productType == ProductType.ptRegular))
                                {
                                    break;
                                }

                                result += String.Format(
                                    "\n\t\t\t<suboption id=\"{0}\">" +
                                    "\n\t\t\t\t<garums_mm>{1}</garums_mm>" +
                                    "\n\t\t\t\t<platums_mm>{2}</platums_mm>",
                                    counter, o2, o3
                                );

                                if (product.productType == ProductType.ptRegular)
                                {
                                    // add data
                                    result += String.Format(
                                        "\n\t\t\t\t<augstums_mm>{0}</augstums_mm>" +
                                        "\n\t\t\t\t<price_eur_no_vat>{1}</price_eur_no_vat>",
                                        o4, o5
                                    );
                                }
                                else
                                if (product.productType == ProductType.ptBed)
                                {
                                    // add data
                                    result += String.Format(
                                        "\n\t\t\t\t<price_eur_no_vat>{0}</price_eur_no_vat>",
                                        o4
                                    );
                                }
                                else
                                {
                                    throw new NotImplementedException();
                                }

                                result += "\n\t\t\t</suboption>";

                                row++;
                            }
                            while (true);

                            result += "\n\t\t</suboptions>";
                            result += "\n\t</product>";

                            // Step back, because row will be increased on next step
                            if ((!String.IsNullOrEmpty(ExcelFunctions.GetString(worksheet, row, 1)) && counter > 1))
                            {
                                row--;
                            }
                        }
                    }
                    while (row < product.end);
                }

                row++;
            }
            while (row < product.end);

            return result;
        }

        /// <summary>
        ///     Collection information about available products on worksheet
        /// </summary>
        /// <param name="worksheet">
        ///     Worksheet for analyze
        /// </param>
        public static List<ProductBlock> AnalyzeWorksheet(ExcelObject.Worksheet worksheet)
        {
            List<ProductBlock> result = new List<ProductBlock>();

            int row = 0;

            int maxRow = worksheet.Cells[worksheet.Rows.Count, 1].End(ExcelObject.XlDirection.xlUp).Row + 10;

            ProductBlock product = null;

            do
            {
                row++;

                // First match
                ProductType isMatch = IsNameInThisRow(worksheet, row);
                if (isMatch != ProductType.ptUnknown)
                {
                    // Second match
                    if (IsHeaderInThisRow(worksheet, row + 1))
                    {
                        // Set end if needed
                        if (product != null)
                        {
                            product.end = row - 1;

                            result.Add(product);

                            product = null;
                        }

                        // MaxRow пока не знаю иначе
                        product = new ProductBlock
                        {
                            productType = isMatch,
                            worksheet = worksheet.Index,
                            worksheetName = worksheet.Name,
                            name = ExcelFunctions.GetString(worksheet, row, isMatch == ProductType.ptRegular ? 6 : 5),
                            start = row,
                            end = maxRow
                        };
                    }

                    row += 2;
                }
            }
            while (row < maxRow);

            // Last one. Could be Null if no product on worksheet. Shouldn't be 0 product on worksheet, but who knows
            if (product != null)
            {
                result.Add(product);
            }

            return result;
        }

        #region Style specific
        /// <summary>
        ///     Validate name
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="row"></param>
        /// <returns>
        /// </returns>
        private static ProductType IsNameInThisRow(ExcelObject.Worksheet worksheet, int row)
        {
            const int CONST_WORKSHEET_PRODUCT_TITLE_COL_ID = 6;

            ProductType result = ProductType.ptUnknown;

            if (_IsNameInThisRow(worksheet, row, CONST_WORKSHEET_PRODUCT_TITLE_COL_ID))
            {
                result = ProductType.ptRegular;
            }
            else
            if (_IsNameInThisRow(worksheet, row, CONST_WORKSHEET_PRODUCT_TITLE_COL_ID - 1))
            {
                result = ProductType.ptBed;
            }

            return result;
        }

        /// <summary>
        ///     Support Name in Col #5 or in #6.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        private static bool _IsNameInThisRow(ExcelObject.Worksheet worksheet, int row, int col)
        {
            bool result = true;

            // This cell has value?
            result = result && !String.IsNullOrEmpty(ExcelFunctions.GetString(worksheet, row, col));

            // "If" for optimization
            if (result)
            {
                // first celsl has empty value
                for (int x = 1; x < col; x++)
                {
                    result = result && String.IsNullOrEmpty(ExcelFunctions.GetString(worksheet, row, x));

                    // "If" for optimization
                    if (!result)
                    {
                        break;
                    }
                }
            }

            return result;
        }

        /// <summary>
        ///     Validate header for product table
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="row"></param>
        /// <returns></returns>
        public static bool IsHeaderInThisRow(ExcelObject.Worksheet worksheet, int row)
        {
            bool result = true;

            string[] headers6 = new string[]
            {
                "№", "Nosaukums", "Garums\n(mm)", "Platums\n(mm)", "Augstums\n(mm)", "EUR bez PVN"
            };

            string[] headers5 = new string[]
            {
                "№", "Nosaukums", "Garums\n(mm)", "Platums\n(mm)", "EUR bez PVN"
            };

            result = _IsHeaderInThisRow(worksheet, row, headers6);

            if (!result)
            {
                result = _IsHeaderInThisRow(worksheet, row, headers5);
            }

            return result;
        }

        private static bool _IsHeaderInThisRow(ExcelObject.Worksheet worksheet, int row, string[] headers)
        {
            bool result = true;

            for (int x = 0; x < headers.Count(); x++)
            {
                if (ExcelFunctions.GetString(worksheet, row, x + 1) != headers[x])
                {
                    result = false;
                    break;
                }
            }

            return result;
        }
        #endregion Style specific
    }
}
