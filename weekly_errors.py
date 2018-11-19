#! /usr/bin/python


# Import the modules that are required for the script
import sys
from weekly_errors_utils import *


if __name__ == '__main__':
	
	######################################
	# IMPORT AND CHECK INTEGREITY (Douglas)
	######################################
	
	# Starting Douglas Error Report
	print('##########################################')
	print('Commencing Douglas error report routine...\n')

	# Import the marked up error file
	file_path_markup = './douglas/error_markup_douglas.xlsx'
	print('Importing previous weekly error report...')
	wb_markup = import_excel(file_path_markup)
	
	# Check the integrity of the previous week file
	if file_integrity_check(wb_markup, 'douglas') == False:
		sys.exit()

	# Ensure that the markup contains no blank comments
	print('Checking markup does not contain blank comments...')
	if markup_complete_check(wb_markup) == False:
		sys.exit()

	# Import the new error file
	file_path_new = './douglas/error_new_douglas.xlsx'
	print('Importing new weekly error report...')
	wb_new = import_excel(file_path_new)

	# Check the integrity of the new file
	if file_integrity_check(wb_new, 'douglas') == False:
		sys.exit()


	######################################
	# RUN ROUTINES (Douglas)
	######################################

	# 2018 TLS-TL0 enrols Routine
	print('Running 2018 TLS-TL0 enrols routine...')
	enrols_2018_TLS_TL0(wb_markup['2018 TLS-TL0 enrols'], wb_new['2018 TLS-TL0 enrols'])
	
	# Act End Dt Error Routine
	print('Running Act End Dt Error routine...')
	act_end_dt_error(wb_markup['Act End Dt Error'], wb_new['Act End Dt Error'])

	# Act St Dt Error Routine
	print('Running Act St Dt Error routine...')
	act_st_dt_error(wb_markup['Act St Dt Error'], wb_new['Act St Dt Error'])

	# dup students Routine
	print('Running dup students routine...')
	dup_students(wb_markup['dup students'], wb_new['dup students'])

	# 2017-18 student dups Routine
	print('Running 2017-18 student dups routine...')
	student_dups_2017_18(wb_markup['2017-18 student dups'], wb_new['2017-18 student dups'])

	# VFH Errors Routine
	print('Running VFH Errors routine...')
	vfh_errors(wb_markup['VFH Errors'], wb_new['VFH Errors'])

	# 2017 enr, 2018 comp Routine
	print('Running 2017 enr, 2018 comp routine...')
	enr_2017_comp_2018(wb_markup['2017 enr, 2018 comp'], wb_new['2017 enr, 2018 comp'])

	# General Errors Routine
	print('Running General Errors routine...')
	general_errors(wb_markup['General Errors'], wb_new['General Errors'])

	# Course intent errors Routine
	print('Running Course intent errors routine...')
	course_intent_errors(wb_markup['Course intent errors'], wb_new['Course intent errors'])

	# Inconsis fee&fund scr-pls chk Routine
	print('Running Inconsis fee&fund scr-pls chk routine...')
	inconsis_fee_fund_scr_pls_chk(wb_markup['Inconsis fee&fund scr-pls chk'], wb_new['Inconsis fee&fund scr-pls chk'])

	# Duplicate SUA Routine
	print('Duplicate SUA routine...')
	duplicate_sua(wb_markup['Duplicate SUA'], wb_new['Duplicate SUA'])

	# superseded unit Routine
	print('superseded unit routine...')
	superseded_unit(wb_markup['superseded unit'], wb_new['superseded unit'])

	# VFH-W chk for participation Routine
	print('Running VFH-W chk for participation routine...')
	vfh_w_chk_for_participation(wb_markup['VFH-W chk for participation'], wb_new['VFH-W chk for participation'])


	######################################
	# COMPLETE JOB (Douglas)
	######################################

	# Save a copy of the final output file
	print('Saving...\n')
	wb_new.save('updated_error_new_douglas.xlsx')

	print('Douglas error report complete.\n')
	
	
	######################################
	# IMPORT AND CHECK INTEGREITY (Margaret)
	######################################
	
	# Starting Margaret Error Report
	print('###########################################')
	print('Commencing Margaret error report routine...\n')

	# Import the marked up error file
	file_path_markup = './margaret/error_markup_margaret.xlsx'
	print('Importing previous weekly error report...')
	wb_markup = import_excel(file_path_markup)
	
	# Check the integrity of the previous week file
	if file_integrity_check(wb_markup, 'margaret') == False:
		sys.exit()

	# Ensure that the markup contains no blank comments
	print('Checking markup does not contain blank comments...')
	if markup_complete_check(wb_markup) == False:
		sys.exit()

	# Import the new error file
	file_path_new = './margaret/error_new_margaret.xlsx'
	print('Importing new weekly error report...')
	wb_new = import_excel(file_path_new)

	# Check the integrity of the new file
	if file_integrity_check(wb_new, 'margaret') == False:
		sys.exit()
	

	######################################
	# RUN ROUTINES (Margaret)
	######################################

	# TCI-not J Routine
	print('Running TCI-not J routine...')
	tci_not_j(wb_markup['TCI-not J'], wb_new['TCI-not J'])

	# K Missing TCI Routine
	print('Running K Missing TCI routine...')
	k_missing_tci(wb_markup['K Missing TCI'], wb_new['K Missing TCI'])

	# Fee and fund incon-pls chk Routine
	print('Running Fee and fund incon-pls chk routine...')
	inconsis_fee_fund_scr_pls_chk(wb_markup['Fee and fund incon-pls chk'], wb_new['Fee and fund incon-pls chk'])


	######################################
	# COMPLETE JOB (Margaret)
	######################################

	# Save a copy of the final output file
	print('Saving...\n')
	wb_new.save('updated_error_new_margaret.xlsx')

	print('Margaret error report complete.\n')
	