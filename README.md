# Dealing with multiple responses

1. command ``mtab``:
``syntax [varlist(default = none)] [, sheet(string) by(string) dir(string) type(string) format(string)]``
Example: ``mtab b51-b515, by(y1) sheet(b5) dir(mtab1.xlsx) type(percentage) format(percent_d2)``

2. program ``multitab``:
``syntax [varlist(default = none)] [, sheet(string) dir(string)]``
Example: ``multitab b51-b515, sheet(b5) dir(multitab.xlsx)``

3. program ``utab``:
``syntax [varlist(default = none)] [, by(string) dir(string) temp(string) type(string) format(string)]``
Example: ``utab dt a8 a9, by(y1) dir(u.xlsx) temp(utab.xls) type(percentage) format(percent_d2)``

4. program ``unitab``:
``syntax [varlist(default = none)] [, dir(string) temp(string)]``
Example: ``unitab a1 a3, dir(unitab.xlsx) temp(temp.xls)``



