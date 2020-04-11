# Program Watchman catches KKs which were shipped or planned to be shipped
# to countries where they are not allowed to sell.

import xlrd
import xlsxwriter

# Tuples with country codes of managed countries.
managed_countries = ("AT", "BE", "BG", "CY", "CZ", "DE", "DK", "EE",
                     "ES", "FI", "FR", "GB", "GR", "HR", "HU", "CH",
                     "IE", "IS", "IT", "LI", "LT", "LV", "MT", "NL",
                     "NO", "PL", "PT", "RO", "SE", "SI", "SK", "LU",
                     "TR", "US", "CA", "BR", "KR", "SG", "MY", "TW",
                     "VN", "ZA", "UK")

# Opening of source document which contains main source data about demands.
source_doc = xlrd.open_workbook("dmds.xlsx")
# Opening of worksheet by name "source_document.sheet_by_name"
# or by index "source_document.sheet_by_index"
source_doc_sheet_1 = source_doc.sheet_by_index(0)

# Creation of new excel file for results.
incorr_dmds = xlsxwriter.Workbook("icorr_dmds.xlsx")
# Opening a worksheet in excel for results.
incorr_dmds_sheet_1 = incorr_dmds.add_worksheet()

# Copy first line of source document into documents with results.
for z in range(source_doc_sheet_1.ncols):
    incorr_dmds_sheet_1.write(0, z, (source_doc_sheet_1.cell(0, z).value))

# Iteration by the countries in source document and search for
# managed countries. When watchmen finds some managed country in demands,
# he write the data about the demand into document with results.
a = 1
for c in range(source_doc_sheet_1.nrows - 1):
    country = source_doc_sheet_1.cell(c + 1, 3).value
    if country in managed_countries:
        for x in range(source_doc_sheet_1.ncols):
            incorr_dmds_sheet_1.write(a, x, (source_doc_sheet_1.cell(c + 1, x).value))
        a += 1

# Close new document with results.
incorr_dmds.close()
