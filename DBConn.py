import cx_oracle

con = cx_Oracle.connect('pythonhol/welcome@127.0.0.1/orcl')
print(con.version)
con.close()

