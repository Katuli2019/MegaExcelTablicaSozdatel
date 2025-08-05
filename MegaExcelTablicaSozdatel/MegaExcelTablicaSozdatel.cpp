#include <iostream>
#include <clocale>
#include <filesystem>
#include <xlnt/xlnt.hpp>
#include <map>
#include <cmath>

xlnt::border::border_property borderProperty;
xlnt::border borderMedium;

struct SortByLengthThenLex {
    bool operator()(const std::string& a, const std::string& b) const {
        if (a.length() == b.length()) {
            return a < b;
        }
        return a.length() < b.length();
    }
};
struct Values
{
    int* marks;
    const char8_t* language = u8"абв";
};
std::map<std::string, Values, SortByLengthThenLex> marksForEachGrade;

void AnalyzeDocument(const std::filesystem::path path);
xlnt::workbook CreateAndFormatOutputExcelFile();
void FillOutputDocument(xlnt::workbook& outputExcel);
void ClearCache();

int main()
{
    std::setlocale(LC_ALL, "ru_RU.UTF-8");
    
    borderProperty.style(xlnt::border_style::thin);
    borderProperty.color(xlnt::color::black());
    
    borderMedium.side(xlnt::border_side::start, borderProperty);
    borderMedium.side(xlnt::border_side::end, borderProperty);
    borderMedium.side(xlnt::border_side::top, borderProperty);
    borderMedium.side(xlnt::border_side::bottom, borderProperty);

    xlnt::workbook outputExcel = CreateAndFormatOutputExcelFile();

    std::string path = ".";
    for (const auto& entry : std::filesystem::directory_iterator(path)) {
        if (entry.path().extension() == ".xlsx") {
            AnalyzeDocument(entry.path());
        }
    }

    FillOutputDocument(outputExcel);
    ClearCache();
    outputExcel.save("Otchet.xlsx");
    return 0;
}

xlnt::font boldTimesNewRoman = xlnt::font().bold(true).size(12).name("Times New Roman");
xlnt::alignment centerAlignment = xlnt::alignment().horizontal(xlnt::horizontal_alignment::center).
    vertical(xlnt::vertical_alignment::center).wrap(true);

void SetUpCellAndFill(xlnt::worksheet& ws, xlnt::cell_reference cell,
    const xlnt::font& font, const xlnt::alignment& alignment)
{
    ws.cell(cell).font(font);
    ws.cell(cell).alignment(alignment);
}
void SetUpCellAndFill(xlnt::worksheet& ws, xlnt::cell_reference cell, double value,
    const xlnt::font& font, const xlnt::alignment& alignment)
{
    ws.cell(cell).value(value);
    SetUpCellAndFill(ws, cell, font, alignment);
}
void SetUpCellAndFill(xlnt::worksheet& ws, xlnt::cell_reference cell, const char8_t* value,
    const xlnt::font& font, const xlnt::alignment& alignment)
{
    ws.cell(cell).value(reinterpret_cast<const char*>(value));
    SetUpCellAndFill(ws, cell, font, alignment);
}
xlnt::workbook CreateAndFormatOutputExcelFile()
{
    xlnt::workbook outputExcel;
    xlnt::worksheet ws = outputExcel.active_sheet();
    ws.title("Otchet");

    for (int i = 1; i <= 12; i++)
    {
        ws.column_properties(xlnt::column_t(i).column_string()).width = 8.11;
    }
    ws.row_properties(3).height = 46.8;

    ws.merge_cells("B1:L1");
    ws.merge_cells("B2:B3");
    ws.merge_cells("C2:C3");
    ws.merge_cells("D2:D3");
    ws.merge_cells("E2:E3");
    ws.merge_cells("F2:J2");
    ws.merge_cells("K2:K3");
    ws.merge_cells("L2:L3");


    SetUpCellAndFill(ws, "B1", u8"Отчет за год  учебного года преподавателя", boldTimesNewRoman, centerAlignment);
    SetUpCellAndFill(ws, "B2", u8"№", boldTimesNewRoman, centerAlignment);
    SetUpCellAndFill(ws, "C2", u8"Класс", boldTimesNewRoman, centerAlignment);
    SetUpCellAndFill(ws, "D2", u8"Язык обучения", boldTimesNewRoman, centerAlignment);
    SetUpCellAndFill(ws, "E2", u8"Количество детей", boldTimesNewRoman, centerAlignment);
    SetUpCellAndFill(ws, "F2", u8"Успеваемость", boldTimesNewRoman, centerAlignment);

    SetUpCellAndFill(ws, "F3", 5, boldTimesNewRoman, centerAlignment);
    SetUpCellAndFill(ws, "G3", 4, boldTimesNewRoman, centerAlignment);
    SetUpCellAndFill(ws, "H3", 3, boldTimesNewRoman, centerAlignment);
    SetUpCellAndFill(ws, "I3", 2, boldTimesNewRoman, centerAlignment);
    SetUpCellAndFill(ws, "J3", u8"не аттестован", boldTimesNewRoman, centerAlignment);

    SetUpCellAndFill(ws, "K2", u8"Успеваемость\n%", boldTimesNewRoman, centerAlignment);
    SetUpCellAndFill(ws, "L2", u8"Качество\n%", boldTimesNewRoman, centerAlignment);

    for (int i = 2; i <= 12; i++)
    {
        for (int j = 2; j <= 3; j++)
        {
			ws.cell(i, j).border(borderMedium);
        }
    }

    return outputExcel;
}

void AnalyzeDocument(std::filesystem::path path)
{
    try
    {
        xlnt::workbook inputWorkbook;
        inputWorkbook.load(path);

        bool isTargetDocument = false;
        int marksColumn = 0;
        for (int i = 1; i < 50; i++)
        {
            if (inputWorkbook.active_sheet().cell(i, 8).to_string() == reinterpret_cast<const char*>(u8"Оценка"))
            {
                marksColumn = i;
                isTargetDocument = true;
                break;
            }
        }
        if (!isTargetDocument)
            return;

        std::string tempString = inputWorkbook.active_sheet().cell(3, 10).to_string();
        std::string grade = tempString.substr(0, tempString.find("("));

		std::string language = tempString.substr(tempString.find("("));
        if (language == reinterpret_cast<const char*>(u8"(Русский язык)"))
        {
            marksForEachGrade[grade].language = u8"рус";
        }
        else if (language == reinterpret_cast<const char*>(u8"(Казахский язык)"))
        {
			marksForEachGrade[grade].language = u8"каз";
        }

        marksForEachGrade[grade].marks = new int[5];
        for (int i = 0; i < 5; i++)
        {
            marksForEachGrade[grade].marks[i] = 0;
		}

        for (int i = 10;; i++)
        {
			std::string potentialMark = inputWorkbook.active_sheet().cell(marksColumn, i).to_string();
            if (potentialMark.empty())
            {
                break;
            }
            else
            {
                if (potentialMark == reinterpret_cast<const char*>(u8"Н/а"))
                {
                    marksForEachGrade[grade].marks[0]++;
                }
                else
                {
					marksForEachGrade[grade].marks[std::stoi(potentialMark) - 1]++;
                }
            }
        }
    }
    catch (const std::exception& e)
    {
        std::cout << "Error opening: " << path.string() << std::endl;
    }
}

void FillOutputDocument(xlnt::workbook& outputExcel)
{
    xlnt::worksheet ws = outputExcel.active_sheet();

    int totalStudents = 0;
	int totalMarks[5] = { 0, 0, 0, 0, 0 };
	double totalAchievement = 0.0;
	double totalQuality = 0.0;

    int i = 0;
    for (const auto& [grade, values] : marksForEachGrade)
    {
        i++;

        for (int j = 2; j <= 12; j++)
        {
            ws.cell(j, i + 3).border(borderMedium);
		}

		int totalStudentsInGrade = values.marks[0] + values.marks[1] + values.marks[2] + values.marks[3] + values.marks[4];
		totalStudents += totalStudentsInGrade;

		ws.cell(2, i + 3).value(i);
		ws.cell(3, i + 3).value(grade);
        ws.cell(4, i + 3).value(reinterpret_cast<const char*>(values.language));
        ws.cell(5, i + 3).value(totalStudentsInGrade);

        for (int i = 0; i < 5; i++)
        {
            totalMarks[i] += values.marks[i];
		}

        ws.cell(6, i + 3).value(values.marks[4]);
        ws.cell(7, i + 3).value(values.marks[3]);
        ws.cell(8, i + 3).value(values.marks[2]);
        ws.cell(9, i + 3).value(values.marks[1]);
        ws.cell(10, i + 3).value(values.marks[0]);

        double achievement = static_cast<double>(values.marks[4] + values.marks[3] + values.marks[2]) * 100 / totalStudentsInGrade;
		totalAchievement += achievement;
        ws.cell(11, i + 3).value(std::round(achievement));

		double quality = static_cast<double>(values.marks[4] + values.marks[3]) * 100 / totalStudentsInGrade;
		totalQuality += quality;
		ws.cell(12, i + 3).value(std::round(quality));
    }

	totalAchievement /= i;
	totalQuality /= i;

    for (int j = 2; j <= 12; j++)
    {
        ws.cell(j, i + 4).border(borderMedium);
    }
    SetUpCellAndFill(ws, xlnt::cell_reference(2, i + 4), u8"Итого:", boldTimesNewRoman, centerAlignment);
	ws.cell(5, i + 4).value(totalStudents);
    ws.cell(6, i + 4).value(totalMarks[4]);
    ws.cell(7, i + 4).value(totalMarks[3]);
    ws.cell(8, i + 4).value(totalMarks[2]);
    ws.cell(9, i + 4).value(totalMarks[1]);
    ws.cell(10, i + 4).value(totalMarks[0]);
    ws.cell(11, i + 4).value(std::round(totalAchievement));
	ws.cell(12, i + 4).value(std::round(totalQuality));

	ws.merge_cells(xlnt::range_reference(5, i + 7, 5, i + 8));
	ws.merge_cells(xlnt::range_reference(6, i + 7, 6, i + 8));
    for (int j = 5; j <= 6; j++)
    {
        ws.cell(j, i + 7).border(borderMedium);
        ws.cell(j, i + 8).border(borderMedium);
    }

    SetUpCellAndFill(ws, xlnt::cell_reference(5, i + 7), u8"Качество знаний в разрезе классов", 
        xlnt::font().bold(true).size(6).name("Times New Roman"), centerAlignment);
	SetUpCellAndFill(ws, xlnt::cell_reference(6, i + 7), u8"Качество", boldTimesNewRoman, centerAlignment);

    i = i + 9;
    std::string currentGrade;
    double averageQualityPerGrade = 0;
	int gradeCount = 0;
	double totalAverageQuality = 0;
	int totalGradeCount = 0;
    for (const auto& [grade, values] : marksForEachGrade)
    {
        std::string gradeString;
        for (const auto& c : grade)
        {
            if (isdigit(c))
            {
                gradeString += c;
            }
		}

        if (!currentGrade.empty() && currentGrade != gradeString)
        {
			averageQualityPerGrade /= gradeCount;
            totalAverageQuality += averageQualityPerGrade;

            for (int j = 5; j <= 6; j++)
            {
                ws.cell(j, i).border(borderMedium);
            }
			ws.cell(6, i).value(std::round(averageQualityPerGrade));
			ws.cell(5, i).value(currentGrade + reinterpret_cast<const char*>(u8" класс"));

			totalGradeCount++;
			averageQualityPerGrade = 0;
			gradeCount = 0;
            currentGrade.clear();
            i++;
        }

        currentGrade = gradeString;
        averageQualityPerGrade += static_cast<double>(values.marks[4] + values.marks[3]) * 100 / (values.marks[4] + values.marks[3] + values.marks[2] + values.marks[1] + values.marks[0]);
        gradeCount++;
	}

    averageQualityPerGrade /= gradeCount;
    totalAverageQuality += averageQualityPerGrade;

    for (int j = 5; j <= 6; j++)
    {
        ws.cell(j, i).border(borderMedium);
    }
    ws.cell(6, i).value(std::round(averageQualityPerGrade));
    ws.cell(5, i).value(currentGrade + reinterpret_cast<const char*>(u8" класс"));

    totalGradeCount++;
    i++;

    for (int j = 5; j <= 6; j++)
    {
        ws.cell(j, i).border(borderMedium);
    }
    totalAverageQuality /= totalGradeCount;
    ws.cell(6, i).value(std::round(totalAverageQuality));
    SetUpCellAndFill(ws, xlnt::cell_reference(5, i), u8"срзнач", boldTimesNewRoman, centerAlignment);
}

void ClearCache()
{
    for (auto& [grade, values] : marksForEachGrade)
    {
        delete[] values.marks;
		//delete[] values.language;
    }
    marksForEachGrade.clear();
}