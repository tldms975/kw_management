import openpyxl



def main():
	print("OPEN FILE...")
	remove_list_path = './불용신청 목록.xlsx'
	remove_list_book = openpyxl.load_workbook(remove_list_path)
	remove_list_sheet = remove_list_book['sheet1']

	target_list_path = './복사본 재물조사 대상 물품.xlsx'
	target_list_book = openpyxl.load_workbook(target_list_path)
	target_list_sheet = target_list_book['sheet1']
	print("OPENED!")
	for remove_row in remove_list_sheet:
		remove_id_cell = remove_row[0]
		print("Start Found: "+str(remove_id_cell.value))
		for target_row in target_list_sheet:
			target_id_cell = target_row[0]
			if target_id_cell.value == remove_id_cell.value:
				print("I Found:"+str(target_id_cell.value)+" / DELETED")
				target_list_sheet.delete_rows(target_row[0].row,1)
				break
	target_list_book.save('./최종본.xlsx')
	pass




if __name__ == '__main__':
	main()