import fs from "fs"
import path from "path"

import XLSX from "xlsx"

console.error = function () {}

const dict = [
	"Январь",
	"Февраль",
	"Март",
	"Апрель",
	"Май",
	"Июнь",
	"Июль",
	"Август",
	"Сентябрь",
	"Октябрь",
	"Ноябрь",
	"Декабрь",
]
interface IJsonOutput {
	[name: string]: string
}

;(function CountXlsx(): void {
	try {
		const sumFinalArray: IJsonOutput[] = []
		fs.readdirSync("./tables", {
			encoding: "utf8",
		}).forEach((file) => {
			if (path.extname(file).toLowerCase() !== ".xlsx") return
			const JsonTable = parseExcel("./tables/" + file)
			let concatArray: number[] = []
			let november = 0
			for (let month = 1; month <= 12; month++) {
				const filtered = JsonTable.filter(
					(date) => Number(date["дата"].split(".")[1]) == month
				)

				let countArr = [0, 0, 0, 0, 0]
				filtered.map((value) => {
					countArr[0] += value["всего кол-во"]
					countArr[1] += value["БК"]
					countArr[2] += value["ТК"]
					countArr[3] += value["СОЦ.К."]
					countArr[4] += value["ТРОЙКА"]

					if (month === 11) {
						const day = Number(value["дата"].split(".")[0])
						const week = value["день недели"]

						if (
							(day > 6 || day < 4) &&
							week !== "вс" &&
							week !== "сб"
						) {
							november += value["всего кол-во"]
						}
					}
				})
				concatArray = concatArray.concat(countArr)
			}
			const JsonOutput: IJsonOutput = {}
			JsonOutput["Файл"] = path.parse(file).name
			JsonOutput["Поток"] = String(november)
			for (let i = 0; i < concatArray.length; i++) {
				JsonOutput[dict[~~(i / 5)] + ((i % 5) + 1)] = String(
					concatArray[i]
				)
			}

			sumFinalArray.push(JsonOutput)
			console.info(
				"Файл '" + path.parse(file).name + "' обработан успешно!"
			)
		})
		writeExcel("./output.xlsx", sumFinalArray)
	} catch (e: any) {
		console.warn("Произошла ошибка при обработке файла")
		throw new Error(e)
	}
})()

function parseExcel(filePath: string): any[] {
	const workBook = XLSX.readFile(filePath)

	let name = workBook.SheetNames[0]

	return XLSX.utils.sheet_to_json(workBook.Sheets[name])
}
function writeExcel(
	filePath: string,
	list: any[],
	sheetName = "Выходные данные"
) {
	const workBook = XLSX.utils.book_new()

	let objectMaxLength: number[] = []

	list.map((jsonData) => {
		Object.entries(jsonData).map(([a, v], idx) => {
			const max = 100
			let columnHeader = a
			let columnValue = v as string
			let columnWitdh =
				columnHeader.length > columnValue.length
					? columnHeader.length
					: columnValue.length
			if (columnWitdh > max) columnWitdh = max
			objectMaxLength[idx] =
				objectMaxLength[idx] >= columnWitdh
					? objectMaxLength[idx]
					: columnWitdh + 1
		})
	})

	const wscols = objectMaxLength.map((w: number) => ({ width: w }))
	XLSX.utils.book_append_sheet(
		workBook,
		XLSX.utils.json_to_sheet(list),
		sheetName
	)

	workBook.Sheets[sheetName]["!cols"] = wscols

	XLSX.writeFile(workBook, filePath)
}
