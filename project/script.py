from helper import *

def main():
	## tests
	row = ["asdf","asdf","asdf","asdf"]
	print("asdf1")
	append_list_to_spreadsheet(row)
	test_output_rows=pull_all_output_rows()
	print(test_output_rows)

if __name__ == '__main__':
    main()
    clear_range("A","2", "F")