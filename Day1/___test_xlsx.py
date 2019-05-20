
#############
### Send mail : echo " Current ICQA VAlidation Excel " | mailx -a /var/SP/DUMP/convers1/kuldeep/python_prac/DR3_ICQA_Validations_Staging.xlsx -s "ICQA DR3 VAlidations" midhilem@amdocs.com,kukumar@amdocs.com
###########

import xlsxwriter as xls
import collections
import os
import csv
import cx_Oracle

# Create an new Excel file and add a worksheet.

workbook = xls.Workbook('/var/SP/DUMP/convers1/DR5_GOLDEN_GATE/SESSION/bin/PRD_ICQA_Validations_Staging.xlsx')

#############################################
#### Connect to ORacle DB ##########
#############################################

a_db_user = 'CNVABP9'
c_db_user = 'SA9'
o_db_user = 'CNVOMS9'
db_pass = 'mig_unify10'
a_db_inst = 'migabp'
c_db_inst = 'migcrm'
o_db_inst = 'migoms'

##################################################
##### Sheet :: Summary
#################################################

#Connect to ABP
db_user = a_db_user
db_inst = a_db_inst

db_conn_str = db_user + '/' + db_pass + '@' + db_inst
a_con = cx_Oracle.connect(db_conn_str)
c = a_con.cursor()

query_sum1 = '''
SELECT 'AREA',
  'CHECKS COUNT',
  'Source IC Results',
  'Target IC Results',
  'Relative DB Results Pass %',
  'Target New Fields Validations Checks Count',
  'Target New Fields Validations Result',
  'Comments'
FROM DUAL
'''
worksheet = workbook.add_worksheet('Summary')

c.execute(query_sum1)
# this will start from Row = 5
row = 4
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column+1, result[column])
        row += 1


worksheet.write(5, 1, 'AR')
worksheet.write(6, 1, 'INV')
worksheet.write(7, 1, 'SRM')
worksheet.write(8, 1, 'CM')
worksheet.write(9, 1, 'OMS')
worksheet.write(10, 1, 'CRM')
# Close Connection : ABP

print "Writing :: Summary for AR "

query_sum_ar3 = '''
select count_check from (
SELECT DECODE(CHECK_TYPE,'OLD_FIELDS','Checks Count','NEW_FIELDS','Target New Fields Validations Checks Count') DETAILS,
  COUNT(CHECK_TYPE) COUNT_CHECK ,
  ROUND(AVG(results),10) AVG_RESULT
FROM
  (SELECT
    CASE
      WHEN DESCRIPTION LIKE 'Validation of new fields 10.2 table%'
      THEN 'NEW_FIELDS'
      WHEN DESCRIPTION NOT LIKE 'Validation of new fields 10.2 table%'
      THEN 'OLD_FIELDS'
    END AS CHECK_TYPE,
    RESULTS
  FROM TAR_CHECKS
  )
GROUP BY CHECK_TYPE )
where details = 'Checks Count'
'''

query_sum_ar4 = '''
SELECT SRC_IC_RESULTS from (
SELECT CHECK_TYPE ,
  COUNT(CHECK_TYPE) ,
  ROUND(AVG(results),10) SRC_IC_RESULTS
FROM
  (SELECT
    CASE
      WHEN DESCRIPTION LIKE 'Validation of new fields 10.2 table%'
      THEN 'NEW_FIELDS'
      WHEN DESCRIPTION NOT LIKE 'Validation of new fields 10.2 table%'
      THEN 'OLD_FIELDS'
    END AS CHECK_TYPE,
    RESULTS
  FROM SAR_CHECKS@CONN_DB_LINK_SRC_ABP.PROD.NL
  )
GROUP BY CHECK_TYPE )
where CHECK_TYPE = 'OLD_FIELDS'
'''
query_sum_ar5 = '''
select AVG_RESULT from (
SELECT DECODE(CHECK_TYPE,'OLD_FIELDS','Checks Count','NEW_FIELDS','Target New Fields Validations Checks Count') DETAILS,
  COUNT(CHECK_TYPE) COUNT_CHECK ,
  ROUND(AVG(results),10) AVG_RESULT
FROM
  (SELECT
    CASE
      WHEN DESCRIPTION LIKE 'Validation of new fields 10.2 table%'
      THEN 'NEW_FIELDS'
      WHEN DESCRIPTION NOT LIKE 'Validation of new fields 10.2 table%'
      THEN 'OLD_FIELDS'
    END AS CHECK_TYPE,
    RESULTS
  FROM TAR_CHECKS
  )
GROUP BY CHECK_TYPE )
where details = 'Checks Count'
'''

query_sum_ar7 = '''SELECT COUNT_CHECK
FROM
  (SELECT DECODE(CHECK_TYPE,'OLD_FIELDS','Checks Count','NEW_FIELDS','Target New Fields Validations Checks Count') DETAILS,
    COUNT(CHECK_TYPE)  COUNT_CHECK,
    ROUND(AVG(results),10) AVG_RESULT
  FROM
    (SELECT
      CASE
        WHEN DESCRIPTION LIKE 'Validation of new fields 10.2 table%'
        THEN 'NEW_FIELDS'
        WHEN DESCRIPTION NOT LIKE 'Validation of new fields 10.2 table%'
        THEN 'OLD_FIELDS'
      END AS CHECK_TYPE,
      RESULTS
    FROM TAR_CHECKS
    )
  GROUP BY CHECK_TYPE
  )
WHERE details = 'Target New Fields Validations Checks Count'
'''

query_sum_ar8 = '''SELECT AVG_RESULT
FROM
  (SELECT DECODE(CHECK_TYPE,'OLD_FIELDS','Checks Count','NEW_FIELDS','Target New Fields Validations Checks Count') DETAILS,
    COUNT(CHECK_TYPE)  COUNT_CHECK,
    ROUND(AVG(results),10) AVG_RESULT
  FROM
    (SELECT
      CASE
        WHEN DESCRIPTION LIKE 'Validation of new fields 10.2 table%'
        THEN 'NEW_FIELDS'
        WHEN DESCRIPTION NOT LIKE 'Validation of new fields 10.2 table%'
        THEN 'OLD_FIELDS'
      END AS CHECK_TYPE,
      RESULTS
    FROM TAR_CHECKS
    )
  GROUP BY CHECK_TYPE
  )
WHERE details = 'Target New Fields Validations Checks Count'
'''


c.execute(query_sum_ar3)
row = 5
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column+2, result[column])
        row += 1

c.execute(query_sum_ar4)
row = 5
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column+3, result[column])
        row += 1

c.execute(query_sum_ar5)
row = 5
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column+4, result[column])
        row += 1

c.execute(query_sum_ar7)
row = 5
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column+6, result[column])
        row += 1

c.execute(query_sum_ar8)
row = 5
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column+7, result[column])
        row += 1


print "Writing :: Summary for INV "

query_sum_inv3 = '''
select count_check from (
SELECT DECODE(CHECK_TYPE,'OLD_FIELDS','Checks Count','NEW_FIELDS','Target New Fields Validations Checks Count') DETAILS,
  COUNT(CHECK_TYPE) COUNT_CHECK ,
  ROUND(AVG(results),10) AVG_RESULT
FROM
  (SELECT
    CASE
      WHEN DESCRIPTION LIKE 'Validation of new fields 10.2 table%'
      THEN 'NEW_FIELDS'
      WHEN DESCRIPTION NOT LIKE 'Validation of new fields 10.2 table%'
      THEN 'OLD_FIELDS'
    END AS CHECK_TYPE,
    RESULTS
  FROM TIN_CHECKS
  )
GROUP BY CHECK_TYPE )
where details = 'Checks Count'
'''

query_sum_inv4 = '''
SELECT SRC_IC_RESULTS from (
SELECT CHECK_TYPE ,
  COUNT(CHECK_TYPE) ,
  ROUND(AVG(results),10) SRC_IC_RESULTS
FROM
  (SELECT
    CASE
      WHEN DESCRIPTION LIKE 'Validation of new fields 10.2 table%'
      THEN 'NEW_FIELDS'
      WHEN DESCRIPTION NOT LIKE 'Validation of new fields 10.2 table%'
      THEN 'OLD_FIELDS'
    END AS CHECK_TYPE,
    RESULTS
  FROM SIN_CHECKS@CONN_DB_LINK_SRC_ABP.PROD.NL
  )
GROUP BY CHECK_TYPE )
where CHECK_TYPE = 'OLD_FIELDS'
'''
query_sum_inv5 = '''
select AVG_RESULT from (
SELECT DECODE(CHECK_TYPE,'OLD_FIELDS','Checks Count','NEW_FIELDS','Target New Fields Validations Checks Count') DETAILS,
  COUNT(CHECK_TYPE) COUNT_CHECK ,
  ROUND(AVG(results),10) AVG_RESULT
FROM
  (SELECT
    CASE
      WHEN DESCRIPTION LIKE 'Validation of new fields 10.2 table%'
      THEN 'NEW_FIELDS'
      WHEN DESCRIPTION NOT LIKE 'Validation of new fields 10.2 table%'
      THEN 'OLD_FIELDS'
    END AS CHECK_TYPE,
    RESULTS
  FROM TIN_CHECKS
  )
GROUP BY CHECK_TYPE )
where details = 'Checks Count'
'''

query_sum_inv7 = '''SELECT COUNT_CHECK
FROM
  (SELECT DECODE(CHECK_TYPE,'OLD_FIELDS','Checks Count','NEW_FIELDS','Target New Fields Validations Checks Count') DETAILS,
    COUNT(CHECK_TYPE)  COUNT_CHECK,
    ROUND(AVG(results),10) AVG_RESULT
  FROM
    (SELECT
      CASE
        WHEN DESCRIPTION LIKE 'Validation of new fields 10.2 table%'
        THEN 'NEW_FIELDS'
        WHEN DESCRIPTION NOT LIKE 'Validation of new fields 10.2 table%'
        THEN 'OLD_FIELDS'
      END AS CHECK_TYPE,
      RESULTS
    FROM TIN_CHECKS
    )
  GROUP BY CHECK_TYPE
  )
WHERE details = 'Target New Fields Validations Checks Count'
'''

query_sum_inv8 = '''SELECT AVG_RESULT
FROM
  (SELECT DECODE(CHECK_TYPE,'OLD_FIELDS','Checks Count','NEW_FIELDS','Target New Fields Validations Checks Count') DETAILS,
    COUNT(CHECK_TYPE)  COUNT_CHECK,
    ROUND(AVG(results),10) AVG_RESULT
  FROM
    (SELECT
      CASE
        WHEN DESCRIPTION LIKE 'Validation of new fields 10.2 table%'
        THEN 'NEW_FIELDS'
        WHEN DESCRIPTION NOT LIKE 'Validation of new fields 10.2 table%'
        THEN 'OLD_FIELDS'
      END AS CHECK_TYPE,
      RESULTS
    FROM TIN_CHECKS
    )
  GROUP BY CHECK_TYPE
  )
WHERE details = 'Target New Fields Validations Checks Count'
'''


c.execute(query_sum_inv3)
row = 6
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column+2, result[column])
        row += 1

c.execute(query_sum_inv4)
row = 6
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column+3, result[column])
        row += 1

c.execute(query_sum_inv5)
row = 6
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column+4, result[column])
        row += 1

c.execute(query_sum_inv7)
row = 6
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column+6, result[column])
        row += 1

c.execute(query_sum_inv8)
row = 6
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column+7, result[column])
        row += 1

print "Writing :: Summary for SRM "

query_sum_srm3 = '''
select count_check from (
SELECT DECODE(CHECK_TYPE,'OLD_FIELDS','Checks Count','NEW_FIELDS','Target New Fields Validations Checks Count') DETAILS,
  COUNT(CHECK_TYPE) COUNT_CHECK ,
  ROUND(AVG(results),10) AVG_RESULT
FROM
  (SELECT
    CASE
      WHEN DESCRIPTION LIKE 'Validation of new fields 10.2 table%'
      THEN 'NEW_FIELDS'
      WHEN DESCRIPTION NOT LIKE 'Validation of new fields 10.2 table%'
      THEN 'OLD_FIELDS'
    END AS CHECK_TYPE,
    RESULTS
  FROM TRM_CHECKS
  )
GROUP BY CHECK_TYPE )
where details = 'Checks Count'
'''

query_sum_srm4 = '''
SELECT SRC_IC_RESULTS from (
SELECT CHECK_TYPE ,
  COUNT(CHECK_TYPE) ,
  ROUND(AVG(results),10) SRC_IC_RESULTS
FROM
  (SELECT
    CASE
      WHEN DESCRIPTION LIKE 'Validation of new fields 10.2 table%'
      THEN 'NEW_FIELDS'
      WHEN DESCRIPTION NOT LIKE 'Validation of new fields 10.2 table%'
      THEN 'OLD_FIELDS'
    END AS CHECK_TYPE,
    RESULTS
  FROM SRM_CHECKS@CONN_DB_LINK_SRC_ABP.PROD.NL
  )
GROUP BY CHECK_TYPE )
where CHECK_TYPE = 'OLD_FIELDS'
'''

query_sum_srm5 = '''
select AVG_RESULT from (
SELECT DECODE(CHECK_TYPE,'OLD_FIELDS','Checks Count','NEW_FIELDS','Target New Fields Validations Checks Count') DETAILS,
  COUNT(CHECK_TYPE) COUNT_CHECK ,
  ROUND(AVG(results),10) AVG_RESULT
FROM
  (SELECT
    CASE
      WHEN DESCRIPTION LIKE 'Validation of new fields 10.2 table%'
      THEN 'NEW_FIELDS'
      WHEN DESCRIPTION NOT LIKE 'Validation of new fields 10.2 table%'
      THEN 'OLD_FIELDS'
    END AS CHECK_TYPE,
    RESULTS
  FROM TRM_CHECKS
  )
GROUP BY CHECK_TYPE )
where details = 'Checks Count'
'''

query_sum_srm7 = '''SELECT COUNT_CHECK
FROM
  (SELECT DECODE(CHECK_TYPE,'OLD_FIELDS','Checks Count','NEW_FIELDS','Target New Fields Validations Checks Count') DETAILS,
    COUNT(CHECK_TYPE)  COUNT_CHECK,
    ROUND(AVG(results),10) AVG_RESULT
  FROM
    (SELECT
      CASE
        WHEN DESCRIPTION LIKE 'Validation of new fields 10.2 table%'
        THEN 'NEW_FIELDS'
        WHEN DESCRIPTION NOT LIKE 'Validation of new fields 10.2 table%'
        THEN 'OLD_FIELDS'
      END AS CHECK_TYPE,
      RESULTS
    FROM TRM_CHECKS
    )
  GROUP BY CHECK_TYPE
  )
WHERE details = 'Target New Fields Validations Checks Count'
'''

query_sum_srm8 = '''SELECT AVG_RESULT
FROM
  (SELECT DECODE(CHECK_TYPE,'OLD_FIELDS','Checks Count','NEW_FIELDS','Target New Fields Validations Checks Count') DETAILS,
    COUNT(CHECK_TYPE)  COUNT_CHECK,
    ROUND(AVG(results),10) AVG_RESULT
  FROM
    (SELECT
      CASE
        WHEN DESCRIPTION LIKE 'Validation of new fields 10.2 table%'
        THEN 'NEW_FIELDS'
        WHEN DESCRIPTION NOT LIKE 'Validation of new fields 10.2 table%'
        THEN 'OLD_FIELDS'
      END AS CHECK_TYPE,
      RESULTS
    FROM TRM_CHECKS
    )
  GROUP BY CHECK_TYPE
  )
WHERE details = 'Target New Fields Validations Checks Count'
'''


c.execute(query_sum_srm3)
row = 7
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column+2, result[column])
        row += 1

c.execute(query_sum_srm4)
row = 7
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column+3, result[column])
        row += 1

c.execute(query_sum_srm5)
row = 7
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column+4, result[column])
        row += 1

c.execute(query_sum_srm7)
row = 7
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column+6, result[column])
        row += 1

c.execute(query_sum_srm8)
row = 7
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column+7, result[column])
        row += 1


print "Writing :: Summary for CM "

query_sum_cm3 = '''
select count_check from (
SELECT DECODE(CHECK_TYPE,'OLD_FIELDS','Checks Count','NEW_FIELDS','Target New Fields Validations Checks Count') DETAILS,
  COUNT(CHECK_TYPE) COUNT_CHECK ,
  ROUND(AVG(results),10) AVG_RESULT
FROM
  (SELECT
    CASE
      WHEN DESCRIPTION LIKE 'Validation of new fields 10.2 table%'
      THEN 'NEW_FIELDS'
      WHEN DESCRIPTION NOT LIKE 'Validation of new fields 10.2 table%'
      THEN 'OLD_FIELDS'
    END AS CHECK_TYPE,
    RESULTS
  FROM TCM_CHECKS
  )
GROUP BY CHECK_TYPE )
where details = 'Checks Count'
'''

query_sum_cm4 = '''
SELECT SRC_IC_RESULTS from (
SELECT CHECK_TYPE ,
  COUNT(CHECK_TYPE) ,
  ROUND(AVG(results),10) SRC_IC_RESULTS
FROM
  (SELECT
    CASE
      WHEN DESCRIPTION LIKE 'Validation of new fields 10.2 table%'
      THEN 'NEW_FIELDS'
      WHEN DESCRIPTION NOT LIKE 'Validation of new fields 10.2 table%'
      THEN 'OLD_FIELDS'
    END AS CHECK_TYPE,
    RESULTS
  FROM SCM_CHECKS@CONN_DB_LINK_SRC_ABP.PROD.NL
  )
GROUP BY CHECK_TYPE )
where CHECK_TYPE = 'OLD_FIELDS'
'''
query_sum_cm5 = '''
select AVG_RESULT from (
SELECT DECODE(CHECK_TYPE,'OLD_FIELDS','Checks Count','NEW_FIELDS','Target New Fields Validations Checks Count') DETAILS,
  COUNT(CHECK_TYPE) COUNT_CHECK ,
  ROUND(AVG(results),10) AVG_RESULT
FROM
  (SELECT
    CASE
      WHEN DESCRIPTION LIKE 'Validation of new fields 10.2 table%'
      THEN 'NEW_FIELDS'
      WHEN DESCRIPTION NOT LIKE 'Validation of new fields 10.2 table%'
      THEN 'OLD_FIELDS'
    END AS CHECK_TYPE,
    RESULTS
  FROM TCM_CHECKS
  )
GROUP BY CHECK_TYPE )
where details = 'Checks Count'
'''

query_sum_cm7 = '''SELECT COUNT_CHECK
FROM
  (SELECT DECODE(CHECK_TYPE,'OLD_FIELDS','Checks Count','NEW_FIELDS','Target New Fields Validations Checks Count') DETAILS,
    COUNT(CHECK_TYPE)  COUNT_CHECK,
    ROUND(AVG(results),10) AVG_RESULT
  FROM
    (SELECT
      CASE
        WHEN DESCRIPTION LIKE 'Validation of new fields 10.2 table%'
        THEN 'NEW_FIELDS'
        WHEN DESCRIPTION NOT LIKE 'Validation of new fields 10.2 table%'
        THEN 'OLD_FIELDS'
      END AS CHECK_TYPE,
      RESULTS
    FROM TCM_CHECKS
    )
  GROUP BY CHECK_TYPE
  )
WHERE details = 'Target New Fields Validations Checks Count'
'''

query_sum_cm8 = '''SELECT AVG_RESULT
FROM
  (SELECT DECODE(CHECK_TYPE,'OLD_FIELDS','Checks Count','NEW_FIELDS','Target New Fields Validations Checks Count') DETAILS,
    COUNT(CHECK_TYPE)  COUNT_CHECK,
    ROUND(AVG(results),10) AVG_RESULT
  FROM
    (SELECT
      CASE
        WHEN DESCRIPTION LIKE 'Validation of new fields 10.2 table%'
        THEN 'NEW_FIELDS'
        WHEN DESCRIPTION NOT LIKE 'Validation of new fields 10.2 table%'
        THEN 'OLD_FIELDS'
      END AS CHECK_TYPE,
      RESULTS
    FROM TCM_CHECKS
    )
  GROUP BY CHECK_TYPE
  )
WHERE details = 'Target New Fields Validations Checks Count'
'''


c.execute(query_sum_cm3)
row = 8
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column+2, result[column])
        row += 1

c.execute(query_sum_cm4)
row = 8
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column+3, result[column])
        row += 1

c.execute(query_sum_cm5)
row = 8
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column+4, result[column])
        row += 1

c.execute(query_sum_cm7)
row = 8
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column+6, result[column])
        row += 1

c.execute(query_sum_cm8)
row = 8
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column+7, result[column])
        row += 1


a_con.close()

################################################
#Connect to OMS
################################################

db_user = o_db_user
db_inst = o_db_inst

db_conn_str = db_user + '/' + db_pass + '@' + db_inst
a_con = cx_Oracle.connect(db_conn_str)
c = a_con.cursor()

print "Writing :: Summary for OMS "

query_sum_oms3 = '''
select count_check from (
SELECT DECODE(CHECK_TYPE,'OLD_FIELDS','Checks Count','NEW_FIELDS','Target New Fields Validations Checks Count') DETAILS,
  COUNT(CHECK_TYPE) COUNT_CHECK ,
  ROUND(AVG(results),10) AVG_RESULT
FROM
  (SELECT
    CASE
      WHEN DESCRIPTION LIKE 'Validation of new fields 10.2 table%'
      THEN 'NEW_FIELDS'
      WHEN DESCRIPTION NOT LIKE 'Validation of new fields 10.2 table%'
      THEN 'OLD_FIELDS'
    END AS CHECK_TYPE,
    RESULTS
  FROM TOM_CHECKS
  )
GROUP BY CHECK_TYPE )
where details = 'Checks Count'
'''

query_sum_oms4 = '''
SELECT SRC_IC_RESULTS from (
SELECT CHECK_TYPE ,
  COUNT(CHECK_TYPE) ,
  ROUND(AVG(results),10) SRC_IC_RESULTS
FROM
  (SELECT
    CASE
      WHEN DESCRIPTION LIKE 'Validation of new fields 10.2 table%'
      THEN 'NEW_FIELDS'
      WHEN DESCRIPTION NOT LIKE 'Validation of new fields 10.2 table%'
      THEN 'OLD_FIELDS'
    END AS CHECK_TYPE,
    RESULTS
  FROM SOM_CHECKS@CONN_DB_LINK_SRC_OMS.PROD.NL
  )
GROUP BY CHECK_TYPE )
where CHECK_TYPE = 'OLD_FIELDS'
'''

query_sum_oms5 = '''
select AVG_RESULT from (
SELECT DECODE(CHECK_TYPE,'OLD_FIELDS','Checks Count','NEW_FIELDS','Target New Fields Validations Checks Count') DETAILS,
  COUNT(CHECK_TYPE) COUNT_CHECK ,
  ROUND(AVG(results),10) AVG_RESULT
FROM
  (SELECT
    CASE
      WHEN DESCRIPTION LIKE 'Validation of new fields 10.2 table%'
      THEN 'NEW_FIELDS'
      WHEN DESCRIPTION NOT LIKE 'Validation of new fields 10.2 table%'
      THEN 'OLD_FIELDS'
    END AS CHECK_TYPE,
    RESULTS
  FROM TOM_CHECKS
  )
GROUP BY CHECK_TYPE )
where details = 'Checks Count'
'''

query_sum_oms7 = '''SELECT COUNT_CHECK
FROM
  (SELECT DECODE(CHECK_TYPE,'OLD_FIELDS','Checks Count','NEW_FIELDS','Target New Fields Validations Checks Count') DETAILS,
    COUNT(CHECK_TYPE)  COUNT_CHECK,
    ROUND(AVG(results),10) AVG_RESULT
  FROM
    (SELECT
      CASE
        WHEN DESCRIPTION LIKE 'Validation of new fields 10.2 table%'
        THEN 'NEW_FIELDS'
        WHEN DESCRIPTION NOT LIKE 'Validation of new fields 10.2 table%'
        THEN 'OLD_FIELDS'
      END AS CHECK_TYPE,
      RESULTS
    FROM TOM_CHECKS
    )
  GROUP BY CHECK_TYPE
  )
WHERE details = 'Target New Fields Validations Checks Count'
'''

query_sum_oms8 = '''SELECT AVG_RESULT
FROM
  (SELECT DECODE(CHECK_TYPE,'OLD_FIELDS','Checks Count','NEW_FIELDS','Target New Fields Validations Checks Count') DETAILS,
    COUNT(CHECK_TYPE)  COUNT_CHECK,
    ROUND(AVG(results),10) AVG_RESULT
  FROM
    (SELECT
      CASE
        WHEN DESCRIPTION LIKE 'Validation of new fields 10.2 table%'
        THEN 'NEW_FIELDS'
        WHEN DESCRIPTION NOT LIKE 'Validation of new fields 10.2 table%'
        THEN 'OLD_FIELDS'
      END AS CHECK_TYPE,
      RESULTS
    FROM TOM_CHECKS
    )
  GROUP BY CHECK_TYPE
  )
WHERE details = 'Target New Fields Validations Checks Count'
'''


c.execute(query_sum_oms3)
row = 9
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column+2, result[column])
        row += 1

c.execute(query_sum_oms4)
row = 9
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column+3, result[column])
        row += 1

c.execute(query_sum_oms5)
row = 9
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column+4, result[column])
        row += 1

c.execute(query_sum_oms7)
row = 9
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column+6, result[column])
        row += 1

c.execute(query_sum_oms8)
row = 9
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column+7, result[column])
        row += 1




a_con.close()

################################################
#Connect to CRM
################################################

db_user = c_db_user
db_inst = c_db_inst

db_conn_str = db_user + '/' + db_pass + '@' + db_inst
a_con = cx_Oracle.connect(db_conn_str)
c = a_con.cursor()

query_sum_crm3 = '''
select count_check from (
SELECT DECODE(CHECK_TYPE,'OLD_FIELDS','Checks Count','NEW_FIELDS','Target New Fields Validations Checks Count') DETAILS,
  COUNT(CHECK_TYPE) COUNT_CHECK ,
  ROUND(AVG(results),10) AVG_RESULT
FROM
  (SELECT
    CASE
      WHEN DESCRIPTION LIKE 'Validation of new fields 10.2 table%'
      THEN 'NEW_FIELDS'
      WHEN DESCRIPTION NOT LIKE 'Validation of new fields 10.2 table%'
      THEN 'OLD_FIELDS'
    END AS CHECK_TYPE,
    RESULTS
  FROM TCR_CHECKS
  )
GROUP BY CHECK_TYPE )
where details = 'Checks Count'
'''

query_sum_crm4 = '''
SELECT SRC_IC_RESULTS from (
SELECT CHECK_TYPE ,
  COUNT(CHECK_TYPE) ,
  ROUND(AVG(results),10) SRC_IC_RESULTS
FROM
  (SELECT
    CASE
      WHEN DESCRIPTION LIKE 'Validation of new fields 10.2 table%'
      THEN 'NEW_FIELDS'
      WHEN DESCRIPTION NOT LIKE 'Validation of new fields 10.2 table%'
      THEN 'OLD_FIELDS'
    END AS CHECK_TYPE,
    RESULTS
  FROM SCR_CHECKS@CONN_DB_LINK_SRC_CRM.PROD.NL
  )
GROUP BY CHECK_TYPE )
where CHECK_TYPE = 'OLD_FIELDS'
'''

query_sum_crm5 = '''
select AVG_RESULT from (
SELECT DECODE(CHECK_TYPE,'OLD_FIELDS','Checks Count','NEW_FIELDS','Target New Fields Validations Checks Count') DETAILS,
  COUNT(CHECK_TYPE) COUNT_CHECK ,
  ROUND(AVG(results),10) AVG_RESULT
FROM
  (SELECT
    CASE
      WHEN DESCRIPTION LIKE 'Validation of new fields 10.2 table%'
      THEN 'NEW_FIELDS'
      WHEN DESCRIPTION NOT LIKE 'Validation of new fields 10.2 table%'
      THEN 'OLD_FIELDS'
    END AS CHECK_TYPE,
    RESULTS
  FROM TCR_CHECKS
  )
GROUP BY CHECK_TYPE )
where details = 'Checks Count'
'''
query_sum_crm7 = '''SELECT COUNT_CHECK
FROM
  (SELECT DECODE(CHECK_TYPE,'OLD_FIELDS','Checks Count','NEW_FIELDS','Target New Fields Validations Checks Count') DETAILS,
    COUNT(CHECK_TYPE)  COUNT_CHECK,
    ROUND(AVG(results),10) AVG_RESULT
  FROM
    (SELECT
      CASE
        WHEN DESCRIPTION LIKE 'Validation of new fields 10.2 table%'
        THEN 'NEW_FIELDS'
        WHEN DESCRIPTION NOT LIKE 'Validation of new fields 10.2 table%'
        THEN 'OLD_FIELDS'
      END AS CHECK_TYPE,
      RESULTS
    FROM TCR_CHECKS
    )
  GROUP BY CHECK_TYPE
  )
WHERE details = 'Target New Fields Validations Checks Count'
'''

query_sum_crm8 = '''SELECT AVG_RESULT
FROM
  (SELECT DECODE(CHECK_TYPE,'OLD_FIELDS','Checks Count','NEW_FIELDS','Target New Fields Validations Checks Count') DETAILS,
    COUNT(CHECK_TYPE)  COUNT_CHECK,
    ROUND(AVG(results),10) AVG_RESULT
  FROM
    (SELECT
      CASE
        WHEN DESCRIPTION LIKE 'Validation of new fields 10.2 table%'
        THEN 'NEW_FIELDS'
        WHEN DESCRIPTION NOT LIKE 'Validation of new fields 10.2 table%'
        THEN 'OLD_FIELDS'
      END AS CHECK_TYPE,
      RESULTS
    FROM TCR_CHECKS
    )
  GROUP BY CHECK_TYPE
  )
WHERE details = 'Target New Fields Validations Checks Count'
'''


c.execute(query_sum_crm3)
row = 10
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column+2, result[column])
        row += 1

c.execute(query_sum_crm4)
row = 10
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column+3, result[column])
        row += 1

c.execute(query_sum_crm5)
row = 10
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column+4, result[column])
        row += 1

c.execute(query_sum_crm7)
row = 10
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column+6, result[column])
        row += 1

c.execute(query_sum_crm8)
row = 10
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column+7, result[column])
        row += 1


a_con.close()



#############################################
#### Run Query #############
#############################################

a_header = '''SELECT 'CHECK_ID',
  'TGT_TOT_CHKD_REC',
  'TGT_NUM_ERR',
  'SRC_TOT_CHKD_REC',
  'SRC_NUM_ERR',
  'TGT_RESULTS',
  'SRC_RESULTS',
  'DESCRIPTION',
  'COMMENTS',
  'QUERY'
FROM DUAL
'''
st_header = '''SELECT 'CHECK_ID',
  'TOT_CHKD_REC',
  'NUM_ERR',
  'RESULTS',
  'DESCRIPTION',
  'QUERY'
FROM DUAL
'''

query_ar = '''SELECT b.check_id check_id,
  b.TOT_CHKD_REC TGT_TOT_CHKD_REC, B.NUM_ERR TGT_NUM_ERR,
  a.TOT_CHKD_REC SRC_TOT_CHKD_REC, a.num_err src_num_err,
  b.results tgt_results,
  a.results src_results,
  B.DESCRIPTION DESCRIPTION,
  NULL  COMMENTS,
  B.QUERY
FROM TAR_CHECKS B
JOIN SAR_CHECKS@CONN_DB_LINK_SRC_ABP.PROD.NL a
ON trim(b.check_id) = trim(a.check_id)
AND B.DESCRIPTION   = a.DESCRIPTION
AND (B.NUM_ERR <> A.NUM_ERR OR B.RESULTS <> A.RESULTS ) '''


query_inv = '''SELECT b.check_id check_id,
  b.TOT_CHKD_REC TGT_TOT_CHKD_REC, B.NUM_ERR TGT_NUM_ERR,
  a.TOT_CHKD_REC SRC_TOT_CHKD_REC, a.num_err src_num_err,
  b.results tgt_results,
  a.results src_results,
  B.DESCRIPTION DESCRIPTION,
  NULL  COMMENTS,
  B.QUERY
FROM TIN_CHECKS B
JOIN SIN_CHECKS@CONN_DB_LINK_SRC_ABP.PROD.NL a
ON trim(b.check_id) = trim(a.check_id)
AND B.DESCRIPTION   = a.DESCRIPTION
AND (B.NUM_ERR <> A.NUM_ERR OR B.RESULTS <> A.RESULTS ) '''


query_rm = '''SELECT b.check_id check_id,
  b.TOT_CHKD_REC TGT_TOT_CHKD_REC, B.NUM_ERR TGT_NUM_ERR,
  a.TOT_CHKD_REC SRC_TOT_CHKD_REC, a.num_err src_num_err,
  b.results tgt_results,
  a.results src_results,
  B.DESCRIPTION DESCRIPTION,
  NULL  COMMENTS,
  B.QUERY
FROM TRM_CHECKS B
JOIN SRM_CHECKS@CONN_DB_LINK_SRC_ABP.PROD.NL a
ON trim(b.check_id) = trim(a.check_id)
AND B.DESCRIPTION   = a.DESCRIPTION
AND (B.NUM_ERR <> A.NUM_ERR OR B.RESULTS <> A.RESULTS ) '''


query_cm = '''SELECT b.check_id check_id,
  b.TOT_CHKD_REC TGT_TOT_CHKD_REC, B.NUM_ERR TGT_NUM_ERR,
  a.TOT_CHKD_REC SRC_TOT_CHKD_REC, a.num_err src_num_err,
  b.results tgt_results,
  a.results src_results,
  B.DESCRIPTION DESCRIPTION,
  NULL  COMMENTS,
  B.QUERY
FROM TCM_CHECKS B
JOIN SCM_CHECKS@CONN_DB_LINK_SRC_ABP.PROD.NL a
ON trim(b.check_id) = trim(a.check_id)
AND B.DESCRIPTION   = a.DESCRIPTION
AND (B.NUM_ERR <> A.NUM_ERR OR B.RESULTS <> A.RESULTS ) '''

query_crm = '''SELECT b.check_id check_id,
  b.TOT_CHKD_REC TGT_TOT_CHKD_REC, B.NUM_ERR TGT_NUM_ERR,
  a.TOT_CHKD_REC SRC_TOT_CHKD_REC, a.num_err src_num_err,
  b.results tgt_results,
  a.results src_results,
  B.DESCRIPTION DESCRIPTION,
  NULL  COMMENTS,
  B.QUERY
FROM TCR_CHECKS@CONN_DB_LINK_TGT_CRM.PROD.NL B
JOIN SCR_CHECKS@CONN_DB_LINK_SRC_CRM.PROD.NL a
ON trim(b.check_id) = trim(a.check_id)
AND B.DESCRIPTION   = a.DESCRIPTION
AND (B.NUM_ERR <> A.NUM_ERR OR B.RESULTS <> A.RESULTS )
'''

query_oms = '''SELECT b.check_id check_id,
  b.TOT_CHKD_REC TGT_TOT_CHKD_REC, B.NUM_ERR TGT_NUM_ERR,
  a.TOT_CHKD_REC SRC_TOT_CHKD_REC, a.num_err src_num_err,
  b.results tgt_results,
  a.results src_results,
  B.DESCRIPTION DESCRIPTION,
  NULL  COMMENTS,
  B.QUERY
FROM TOM_CHECKS@CONN_DB_LINK_TGT_OMS.PROD.NL B
JOIN SOM_CHECKS@CONN_DB_LINK_SRC_OMS.PROD.NL a
ON trim(b.check_id) = trim(a.check_id)
AND B.DESCRIPTION   = a.DESCRIPTION
AND (B.NUM_ERR <> A.NUM_ERR OR B.RESULTS <> A.RESULTS )
'''


query_sar = '''SELECT CHECK_ID ,
  TOT_CHKD_REC ,
  NUM_ERR ,
  RESULTS ,
  DESCRIPTION ,
  QUERY
FROM SAR_CHECKS@CONN_DB_LINK_SRC_ABP.PROD.NL
ORDER BY TO_NUMBER (REGEXP_REPLACE (CHECK_ID, '[^0-9]+', ''))
'''

query_sin = '''SELECT CHECK_ID ,
  TOT_CHKD_REC ,
  NUM_ERR ,
  RESULTS ,
  DESCRIPTION ,
  QUERY
FROM SIN_CHECKS@CONN_DB_LINK_SRC_ABP.PROD.NL
ORDER BY TO_NUMBER (REGEXP_REPLACE (CHECK_ID, '[^0-9]+', ''))
'''

query_srm = '''SELECT CHECK_ID ,
  TOT_CHKD_REC ,
  NUM_ERR ,
  RESULTS ,
  DESCRIPTION ,
  QUERY
FROM SRM_CHECKS@CONN_DB_LINK_SRC_ABP.PROD.NL
ORDER BY TO_NUMBER (REGEXP_REPLACE (CHECK_ID, '[^0-9]+', ''))
'''


query_scm = '''SELECT CHECK_ID ,
  TOT_CHKD_REC ,
  NUM_ERR ,
  RESULTS ,
  DESCRIPTION ,
  QUERY
FROM SCM_CHECKS@CONN_DB_LINK_SRC_ABP.PROD.NL
ORDER BY TO_NUMBER (REGEXP_REPLACE (CHECK_ID, '[^0-9]+', ''))
'''

query_scr = '''SELECT CHECK_ID ,
  TOT_CHKD_REC ,
  NUM_ERR ,
  RESULTS ,
  DESCRIPTION ,
  QUERY
FROM SCR_CHECKS@CONN_DB_LINK_SRC_CRM.PROD.NL
ORDER BY TO_NUMBER (REGEXP_REPLACE (CHECK_ID, '[^0-9]+', ''))
'''

query_som = '''SELECT CHECK_ID ,
  TOT_CHKD_REC ,
  NUM_ERR ,
  RESULTS ,
  DESCRIPTION ,
  QUERY
FROM SOM_CHECKS@CONN_DB_LINK_SRC_OMS.PROD.NL
ORDER BY TO_NUMBER (REGEXP_REPLACE (CHECK_ID, '[^0-9]+', ''))
'''


query_tar = '''SELECT CHECK_ID ,
  TOT_CHKD_REC ,
  NUM_ERR ,
  RESULTS ,
  DESCRIPTION ,
  QUERY
FROM TAR_CHECKS
ORDER BY TO_NUMBER (REGEXP_REPLACE (CHECK_ID, '[^0-9]+', ''))
'''

query_tin = '''SELECT CHECK_ID ,
  TOT_CHKD_REC ,
  NUM_ERR ,
  RESULTS ,
  DESCRIPTION ,
  QUERY
FROM TIN_CHECKS
ORDER BY TO_NUMBER (REGEXP_REPLACE (CHECK_ID, '[^0-9]+', ''))
'''

query_trm = '''SELECT CHECK_ID ,
  TOT_CHKD_REC ,
  NUM_ERR ,
  RESULTS ,
  DESCRIPTION ,
  QUERY
FROM TRM_CHECKS
ORDER BY TO_NUMBER (REGEXP_REPLACE (CHECK_ID, '[^0-9]+', ''))
'''


query_tcm = '''SELECT CHECK_ID ,
  TOT_CHKD_REC ,
  NUM_ERR ,
  RESULTS ,
  DESCRIPTION ,
  QUERY
FROM TCM_CHECKS
ORDER BY TO_NUMBER (REGEXP_REPLACE (CHECK_ID, '[^0-9]+', ''))
'''

query_tcr = '''SELECT CHECK_ID ,
  TOT_CHKD_REC ,
  NUM_ERR ,
  RESULTS ,
  DESCRIPTION ,
  QUERY
FROM TCR_CHECKS@CONN_DB_LINK_TGT_CRM.PROD.NL
ORDER BY TO_NUMBER (REGEXP_REPLACE (CHECK_ID, '[^0-9]+', ''))
'''

query_tom = '''SELECT CHECK_ID ,
  TOT_CHKD_REC ,
  NUM_ERR ,
  RESULTS ,
  DESCRIPTION ,
  QUERY
FROM TOM_CHECKS@CONN_DB_LINK_TGT_OMS.PROD.NL
ORDER BY TO_NUMBER (REGEXP_REPLACE (CHECK_ID, '[^0-9]+', ''))
'''



#############################################
###  Write Query Output to Excel ###########
#############################################


##########
## ABP
##########

db_user = a_db_user
db_inst = a_db_inst

db_conn_str = db_user + '/' + db_pass + '@' + db_inst
a_con = cx_Oracle.connect(db_conn_str)
c = a_con.cursor()

# Writing AR Analysis
print "Writing :: AR Analysis "

c.execute(a_header)
worksheet = workbook.add_worksheet('AR Analysis')
row = 0
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column, result[column])
        row += 1

c.execute(query_ar)
row = 1
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column, result[column])
        row += 1

# Writing AR Source
print "Writing :: AR Source "
c.execute(st_header)
worksheet = workbook.add_worksheet('AR Source')
row = 0
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column, result[column])
        row += 1


c.execute(query_sar)
row = 1
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column, result[column])
        row += 1


# Writing AR Target
print "Writing :: AR Target "
c.execute(st_header)
worksheet = workbook.add_worksheet('AR Target')
row = 0
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column, result[column])
        row += 1

c.execute(query_tar)
row = 1
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column, result[column])
        row += 1

# Writing INV Analysis
print "Writing :: INV Analysis "

c.execute(a_header)
worksheet = workbook.add_worksheet('INV Analysis')
row = 0
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column, result[column])
        row += 1

c.execute(query_inv)
row = 1
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column, result[column])
        row += 1

# Writing INV Source
print "Writing :: INV Source "
c.execute(st_header)
worksheet = workbook.add_worksheet('INV Source')
row = 0
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column, result[column])
        row += 1


c.execute(query_sin)
row = 1
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column, result[column])
        row += 1


# Writing INV Target
print "Writing :: INV Target "
c.execute(st_header)
worksheet = workbook.add_worksheet('INV Target')
row = 0
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column, result[column])
        row += 1

c.execute(query_tin)
row = 1
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column, result[column])
        row += 1

# Writing SRM Analysis
print "Writing :: SRM Analysis "

c.execute(a_header)
worksheet = workbook.add_worksheet('SRM Analysis')
row = 0
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column, result[column])
        row += 1

c.execute(query_rm)
row = 1
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column, result[column])
        row += 1

# Writing SRM Source
print "Writing :: SRM Source "
c.execute(st_header)
worksheet = workbook.add_worksheet('SRM Source')
row = 0
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column, result[column])
        row += 1


c.execute(query_srm)
row = 1
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column, result[column])
        row += 1


# Writing SRM Target
print "Writing :: SRM Target "
c.execute(st_header)
worksheet = workbook.add_worksheet('SRM Target')
row = 0
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column, result[column])
        row += 1

c.execute(query_trm)
row = 1
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column, result[column])
        row += 1



# Writing CM Analysis
print "Writing :: CM Analysis "

c.execute(a_header)
worksheet = workbook.add_worksheet('CM Analysis')
row = 0
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column, result[column])
        row += 1

c.execute(query_cm)
row = 1
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column, result[column])
        row += 1

# Writing CM Source
print "Writing :: CM Source "
c.execute(st_header)
worksheet = workbook.add_worksheet('CM Source')
row = 0
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column, result[column])
        row += 1


c.execute(query_scm)
row = 1
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column, result[column])
        row += 1


# Writing CM Target
print "Writing :: CM Target "
c.execute(st_header)
worksheet = workbook.add_worksheet('CM Target')
row = 0
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column, result[column])
        row += 1

c.execute(query_tcm)
row = 1
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column, result[column])
        row += 1

##########
## CRM
##########

# Writing CRM Analysis
print "Writing :: CRM Analysis "

c.execute(a_header)
worksheet = workbook.add_worksheet('CRM Analysis')
row = 0
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column, result[column])
        row += 1

c.execute(query_crm)
row = 1
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column, result[column])
        row += 1

# Writing CRM Source
print "Writing :: CRM Source "
c.execute(st_header)
worksheet = workbook.add_worksheet('CRM Source')
row = 0
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column, result[column])
        row += 1


c.execute(query_scr)
row = 1
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column, result[column])
        row += 1


# Writing CRM Target
print "Writing :: CRM Target "
c.execute(st_header)
worksheet = workbook.add_worksheet('CRM Target')
row = 0
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column, result[column])
        row += 1

c.execute(query_tcr)
row = 1
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column, result[column])
        row += 1


##########
## OMS
##########

# Writing ORM Analysis
print "Writing :: OMS Analysis "

c.execute(a_header)
worksheet = workbook.add_worksheet('OMS Analysis')
row = 0
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column, result[column])
        row += 1

c.execute(query_oms)
row = 1
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column, result[column])
        row += 1



# Writing OMS Source
print "Writing :: OMS Source "
c.execute(st_header)
worksheet = workbook.add_worksheet('OMS Source')
row = 0
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column, result[column])
        row += 1


c.execute(query_som)
row = 1
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column, result[column])
        row += 1


# Writing OMS Target
print "Writing :: OMS Target "
c.execute(st_header)
worksheet = workbook.add_worksheet('OMS Target')
row = 0
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column, result[column])
        row += 1


c.execute(query_tom)
row = 1
for result in c:
        for column in range(len(result)):
                worksheet.write(row, column, result[column])
        row += 1


# Widen the first column to make the text clearer.
#worksheet.set_column('A:A', 20)

# Add a bold format to use to highlight cells.
#bold = workbook.add_format({'bold': True})

# Write some simple text.
#worksheet.write('A1', 'Hello')

# Text with formatting.
#worksheet.write('A2', 'World', bold)

# Write some numbers, with row/column notation.
#worksheet.write(2, 0, 123)
#worksheet.write(3, 0, 123.456)

# Insert an image.
#textfile = open('/var/SP/DUMP/convers1/kuldeep/python_prac/tout.csv')
#row = col = 0
#for line in textfile:
#       worksheet.write(row, col, line.rstrip("\n"))
#       row += 1

#worksheet = workbook.add_worksheet()
#worksheet.write('B5','Kuldeep')


workbook.close()
