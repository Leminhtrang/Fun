import xlwings as xw

import numpy as np



class take_data():

	def __init__(self):

		self.distance_stars = 3
 
		self.distance = 3 # when have two sheet

		self.sheet_write_data = 0

		self.write_range = "A"

		self.take_type = [0] # choose "all" when you want take all sheets

		self.row_quarter = 0

		self.row_name =1

		self.row_title =0

		self.quantity_of_abandoned_goods = 1

		self.separate_name = []

		self.separate_quarters = []

		self.lis_title = []

	def take_data_excel(self,file_data):

		try:

			self.lis_data = []

			self.workbook = xw.Book(file_data)

			if self.take_type == "all":

				self.source_name_sheets = [x.name for x in self.workbook.sheets]

			else:

				self.source_name_sheets = [self.workbook.sheets[x].name for x in self.take_type]
			

			for x in self.source_name_sheets:

				sheet = self.workbook.sheets[x]

				range_address = sheet.api.UsedRange.Address

				data_excel = sheet.range(range_address).value

				if data_excel == None:

					pass

				else:

					self.lis_data.append(np.array(data_excel))

			
			return self.lis_data
		except Exception as e:
			print(f"Error: {e}")
			return []

	def clean_data(self,data_raw):

		lis_data_not_none = []

		

		for data_sheet in data:

			for i in range(1, data_sheet.shape[0]):
				for j in range(data_sheet.shape[1]):
					if data_sheet[i, j] is None:
						data_sheet[i, j] = data_sheet[i - 1, j]

			lis_data_not_none.append(data_sheet[self.quantity_of_abandoned_goods:])

			self.lis_title.append(data_sheet[self.row_title])

		
		for x in lis_data_not_none:

			name = np.unique(x[:, self.row_name])

			quarters = np.unique(x[:, self.row_quarter])
			
			self.separate_name.append(name.tolist())

			self.separate_quarters.append(quarters.tolist())

		

		return lis_data_not_none

	def quarter_cup(self,data):

		quarter_has_been_split = []

		for d,y in zip(data,self.separate_quarters):

			sub_indices = []

			for quarter in y:

				indices = np.where(d[:, self.row_quarter] == quarter)[0]

				sub_indices.append(indices.tolist())

			sub_quarter = []

			for data_quarter in sub_indices:

				sub_quarter.append(d[data_quarter])

			quarter_has_been_split.append(sub_quarter)


		return quarter_has_been_split

	def calculate_total(self,matrix):

		result = []
		
		for sub_matrix in matrix:


			sub_result = []
			for elemen in sub_matrix:

				list_quarter = []
						
				for x in np.unique(elemen[:, self.row_name]):

					
					

					matrix_name = elemen[np.where(elemen ==x)[0]]

					quarter_name = str(matrix_name[:,self.row_quarter][0])

					total = matrix_name[:, 4:].astype(float ).sum(axis=0)

					total = [quarter_name, x+" Total", '', '', *total]

					new_matrix = np.vstack((matrix_name, total))

					list_quarter.append(np.vstack(new_matrix))



				
				positions = np.where(np.char.find(np.vstack(list_quarter)[:, self.row_name].astype(str)," Total") != -1)[0]


				matrix_quarter = np.vstack(list_quarter)[positions]


				quarter_year = str(matrix_quarter[:,self.row_quarter][0])

				total2 = matrix_quarter[:, 4:].astype(float ).sum(axis=0)

				total2 = [quarter_year+" Total", "", '', '', *total2]

				
				new_matrix2 = np.vstack((list_quarter[0], total2))

				sub_result.append(new_matrix2)

			result.append(np.vstack(sub_result))


		return result
				

				

				
	def calculate_grand(self,matrix_calculate):

		result_end = []

		for data in matrix_calculate:



			positions2 = np.where(np.char.find(data[:, self.row_quarter].astype(str)," Total") != -1)[0]

			matrix_grand = data[positions2]
			

			total_grand = matrix_grand[:, 4:].astype(float).sum(axis=0)

			total_grand = ["Grand Total", "", '', '', *total_grand]

			end = np.vstack((data, total_grand))

			result_end.append(end)


		return result_end

	def merge(self,data_end):

		result_merge = []

		for data in data_end:

			sub_result = []

			sub_result = np.empty_like(data, dtype=object)

						# Initialize a variable to keep track of the previous value in the first column
			prev_value = ""

			# Loop through the rows in the data
			for i, row in enumerate(data):
				# Check if the current value in the first column is the same as the previous one
				if row[0] == prev_value:
					sub_result[i, 0] = ""
				else:
					sub_result[i, 0] = row[0]
					prev_value = row[0]

			# Copy the remaining columns as is
			sub_result[:, 1:] = data[:, 1:]

			result_merge.append(sub_result)

		
		return result_merge


	def create_a_results_table(self,result):

		result_table = []

		for data,t in zip(result,self.lis_title):

			#title = ["Project Close Date","SalesRep","Project Name","Forecast Type","Project Amount","Sum Of Amount To Be Billed By Phase"]

			merge_title = np.vstack((t,data))

			
			result_table.append(merge_title)		

		
				
		return result_table

	def write_result(self,file_write,result):


		try:
			pass

			self.workbook2 = xw.Book(file_write)

			
			self.sht_write = self.workbook2.sheets[self.sheet_write_data]


			last_row_cell = self.sht_write.range(self.write_range + str(self.sht_write.cells.last_cell.row)).end('up').row + self.distance_stars



			for x in result:
			
				num_rows, num_cols = x.shape			

				start_cell = self.sht_write.range(self.write_range+str(last_row_cell))

				start_cell.value = x

				end_cell = start_cell.offset(num_rows - 1, num_cols - 1)

				range_address = f"{start_cell.address}:{end_cell.address}"

				self.sht_write.range(range_address).api.EntireColumn.AutoFit()
				for edge in [7, 8, 9, 10, 11, 12]:

					self.sht_write.range(range_address).api.Borders(edge).LineStyle = 1

				
				last_row_cell = self.sht_write.range(self.write_range + str(self.sht_write.cells.last_cell.row)).end('up').row+self.distance

				
			
			
		except Exception as e:
			print(f"Error: {e}")
			return []


				
if __name__ == "__main__":

	file_data = "Table.xlsx"

	file_write = "Table.xlsx"

	

	function = take_data()

	data = function.take_data_excel(file_data)

	unprocessed_data = function.clean_data(data)

	sub_matrix = function.quarter_cup(unprocessed_data)

	total_name = function.calculate_total(sub_matrix)

	grand = function.calculate_grand(total_name)

	matrix_megre = function.merge(grand)

	result = function.create_a_results_table(matrix_megre)

	function.write_result(file_write,result)


	# Fill_in the_results