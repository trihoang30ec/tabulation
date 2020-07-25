* syntax: mtab [varlist], by() sheet() dir() type() format()
/* 
by: cross-tabulation
sheet: sheet name
dir: output file(.xlsx)
type: "percentage" or "frequency" (equivalent to relative frequency table or frequency table, respectively)
format: numeric format such as "number", "percent_d2", etc.
*/
program mtab
version 15
syntax [varlist(default = none)] [, sheet(string) by(string) dir(string) type(string) format(string)]
local byvalues : value label `by'
levelsof `by', local(bylevels)
local bycount : word count `bylevels'
local ++bycount
putexcel set "`dir'", sheet(`sheet') modify
local startcol = 65 /* column A */
foreach i in `bylevels' {
tabm `varlist' if `by' == `i', matcell(frequency) /* ssc install tab_chi */
local ++startcol
mata: st_matrix("percentage", (st_matrix("frequency"):/rowsum(st_matrix("frequency")))) /* create relative frequency matrix "percentage" */
local colname = char(`startcol') /* give a name to a column in Excel */
putexcel `colname'3 = matrix(`type'), nformat(`format') /* put "percentage" or "frequency" matrix */
local colcontent : label `byvalues' `i' 
putexcel `colname'1 = "`colcontent'", left /* add label to Excel columns */
local startcol = `startcol' + `r(c)' - 1
local colname2 = char(`startcol')
putexcel `colname'1:`colname2'1, merge font(ariel, 11, black) txtwrap hcenter
}

local ++startcol
tabm `varlist', matcell(frequency) /* "General" column */
mata: st_matrix("percentage", (st_matrix("frequency"):/rowsum(st_matrix("frequency"))))
local colname = char(`startcol')
putexcel `colname'3 = matrix(`type'), nformat(`format')
putexcel `colname'1 = "Chung", left
local startcol = `startcol' + `r(c)' - 1
local colname2 = char(`startcol')
putexcel `colname'1:`colname2'1, merge font(ariel, 11, black) txtwrap hcenter
local colname = char(`startcol')
local row = 2
foreach rowname in `varlist' {
local ++row
local rowcontent : variable label `rowname'
putexcel A`row' = "`rowcontent'", left
}
foreach varlevels in `varlist' {
labellist `varlevels'
}
local d = 0
local labelvalues 
foreach varlabelvalues in `r(values)' {
local ++d
local labelvalues`d' : word `d' of `r(values)'
local comma ","
local labelvalues `labelvalues' `labelvalues`d'' `comma'
di "`labelvalues'"
}
local labelvalues = strreverse(subinstr(strreverse("`labelvalues'"), ",", "", 1))
di "`labelvalues'"
di "`varlevels'"
mata: labelmatrix = st_vlmap(st_varvaluelabel("`r(varlist)'"), (`labelvalues')) /* get value labels */
mata: valuecolname = J(1, `bycount', labelmatrix) /* create a (1 x n) row vector of the short sequence repeated "bycount" times */
mata: valuecolname
preserve
clear
set obs 1
getmata (var*) = valuecolname /* get string matrix from mata */
export excel using "`dir'", sheet(`sheet') sheetmodify keepcellfmt cell(B2)
restore
putexcel A1:`colname'`row', font(ariel, 11, black)
putexcel A1:`colname'2, border(bottom, medium)
putexcel A`row':`colname'`row', border(bottom, medium)
end

* Example:
mtab b1_01_b-b1_10_b, by(tinh) sheet(b1) dir(frequency_table.xlsx) type(frequency) format(number) /* frequency table */
mtab b2_1-b2_8, by(tinh) sheet(b2) dir(percentage_table.xlsx) type(percentage) format(percent_d2) /* percentage table */








