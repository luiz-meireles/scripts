# -*- coding: utf-8 -*-
import openpyxl
import datetime
from random import randint
import argparse

sheet = openpyxl.Workbook()
active_sheet = sheet.active
active_sheet.title = "Generated fake users"

dr = ['sp', 'rj', 'mg']
turma = ['special-ops', 'team2', 'team3', 'team2']
uni = ['10 - maracatins', '15-geekie']
turno = [ 'Manhã', 'Tarde', 'Noite']

def main(start_id, end_id):
	active_sheet.append([
		u"identificador único",
		u"nome comoleto",
		u"email",
		u"ano",
		u"turma",
		u"turno",
		u"departamento regional",
		u"unidade",
		u"segmento em",
		u"perfil",
		u"status",
		u"idgeekie",
		u"idgeekieuni"
		])

	def get_any(list):
		return list[randint(0, len(list) - 1)]

	output_file_name = u"{}- generated.xlsx".format(
			datetime.datetime.now()
		)

	for i in range(start_id, end_id + 1, 1):
		active_sheet.append([
				str(i),
				u"Nome{}".format(i),
				u"nome{}@teste.com.br".format(i),
				str(randint(1, 9)),
				get_any(turma),
				get_any(turno),
				get_any(dr),
				get_any(uni),
				u"Novo EM",
				u"estudante",
				u"",
				u"",
				u"",
			])

	sheet.save(filename=output_file_name)



if __name__ == "__main__":
	parser = argparse.ArgumentParser(
        description="creates a fake sheet."
    )
	parser.add_argument("--start-id", required=True, dest="start_id")
	parser.add_argument("--range-ids", required=True, dest="range_ids")
	args = parser.parse_args()
	main(int(args.start_id), int(args.range_ids))