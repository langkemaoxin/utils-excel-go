package main

import (
	"flag"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

type SplitConfig struct {
	InputFile      string
	OutputDir      string
	Sheets         []string
	CopyStyles     bool
	CopyFormulas   bool
	AddTimestamp   bool
	PreserveFormat bool
}

type SheetInfo struct {
	Name     string
	RowCount int
	ColCount int
}

type sheetData struct {
	Name     string
	Header   []string
	DataRows [][]string
}

type mergeGroup struct {
	Header []string
	Sheets []sheetData
}

func SplitExcelSheets(inputFile, outputDir string) error {
	return SplitExcelSheetsAdvanced(SplitConfig{
		InputFile:    inputFile,
		OutputDir:    outputDir,
		CopyStyles:   true,
		CopyFormulas: false,
	})
}

func SplitExcelSheetsAdvanced(config SplitConfig) error {
	srcFile, err := excelize.OpenFile(config.InputFile)
	if err != nil {
		return fmt.Errorf("打开文件失败: %w", err)
	}
	defer srcFile.Close()

	sheetsToProcess, err := getSheetsToProcess(srcFile, config.Sheets)
	if err != nil {
		return err
	}

	if err := os.MkdirAll(config.OutputDir, 0755); err != nil {
		return fmt.Errorf("创建输出目录失败: %w", err)
	}

	baseName := strings.TrimSuffix(filepath.Base(config.InputFile), filepath.Ext(config.InputFile))

	for _, sheetName := range sheetsToProcess {
		if err := exportSingleSheet(srcFile, config, baseName, sheetName); err != nil {
			fmt.Printf("导出 Sheet %q 失败: %v\n", sheetName, err)
		}
	}

	return nil
}

func MergeExcelSheetsByHeader(config SplitConfig) error {
	srcFile, err := excelize.OpenFile(config.InputFile)
	if err != nil {
		return fmt.Errorf("打开文件失败: %w", err)
	}
	defer srcFile.Close()

	sheetsToProcess, err := getSheetsToProcess(srcFile, config.Sheets)
	if err != nil {
		return err
	}

	if err := os.MkdirAll(config.OutputDir, 0755); err != nil {
		return fmt.Errorf("创建输出目录失败: %w", err)
	}

	groups := make(map[string]*mergeGroup)
	groupOrder := make([]string, 0, len(sheetsToProcess))

	for _, sheetName := range sheetsToProcess {
		rows, err := srcFile.GetRows(sheetName)
		if err != nil {
			fmt.Printf("读取 Sheet %q 失败: %v\n", sheetName, err)
			continue
		}

		headerIndex, header, ok := extractHeader(rows)
		if !ok {
			fmt.Printf("Sheet %q 没有可识别的表头，已跳过\n", sheetName)
			continue
		}

		key := headerSignature(header)
		group, exists := groups[key]
		if !exists {
			group = &mergeGroup{Header: header}
			groups[key] = group
			groupOrder = append(groupOrder, key)
		}

		group.Sheets = append(group.Sheets, sheetData{
			Name:     sheetName,
			Header:   header,
			DataRows: rows[headerIndex+1:],
		})
	}

	if len(groupOrder) == 0 {
		return fmt.Errorf("没有找到可合并的 Sheet")
	}

	for _, key := range groupOrder {
		group := groups[key]
		if err := exportMergedGroup(srcFile, config, group); err != nil {
			return err
		}
	}

	return nil
}

func exportSingleSheet(srcFile *excelize.File, config SplitConfig, baseName, sheetName string) error {
	dstFile := excelize.NewFile()
	defer dstFile.Close()

	rows, err := srcFile.GetRows(sheetName)
	if err != nil {
		return fmt.Errorf("读取行数据失败: %w", err)
	}

	if len(rows) == 0 {
		return fmt.Errorf("sheet 没有数据")
	}

	for rowIdx, row := range rows {
		for colIdx, cellValue := range row {
			cell, err := excelCellName(rowIdx+1, colIdx+1)
			if err != nil {
				return err
			}

			if config.CopyFormulas {
				if formula, err := srcFile.GetCellFormula(sheetName, cell); err == nil && formula != "" {
					if err := dstFile.SetCellFormula("Sheet1", cell, formula); err != nil {
						return fmt.Errorf("写入公式失败: %w", err)
					}
					continue
				}
			}

			if err := dstFile.SetCellValue("Sheet1", cell, cellValue); err != nil {
				return fmt.Errorf("写入单元格失败: %w", err)
			}
		}
	}

	if config.CopyStyles {
		copyColumnWidths(srcFile, dstFile, sheetName, "Sheet1")
		copyRowHeights(srcFile, dstFile, sheetName, "Sheet1", len(rows))
	}

	outputFileName := generateSplitFileName(baseName, sheetName, config.AddTimestamp)
	outputPath := filepath.Join(config.OutputDir, outputFileName)
	if err := dstFile.SaveAs(outputPath); err != nil {
		return fmt.Errorf("保存文件失败: %w", err)
	}

	fmt.Printf("已导出: %s\n", outputPath)
	return nil
}

func exportMergedGroup(srcFile *excelize.File, config SplitConfig, group *mergeGroup) error {
	if group == nil || len(group.Sheets) == 0 {
		return nil
	}

	dstFile := excelize.NewFile()
	defer dstFile.Close()

	targetSheet := "Sheet1"
	currentRow := 1

	if err := writeRow(dstFile, targetSheet, currentRow, group.Header); err != nil {
		return err
	}
	currentRow++

	for _, sheet := range group.Sheets {
		for _, row := range sheet.DataRows {
			if err := writeRow(dstFile, targetSheet, currentRow, row); err != nil {
				return err
			}
			currentRow++
		}
	}

	if config.CopyStyles {
		firstSheetName := group.Sheets[0].Name
		copyColumnWidths(srcFile, dstFile, firstSheetName, targetSheet)
	}

	outputPath := filepath.Join(config.OutputDir, generateMergedFileName(group.Sheets, config.AddTimestamp))
	if err := dstFile.SaveAs(outputPath); err != nil {
		return fmt.Errorf("保存合并文件失败: %w", err)
	}

	fmt.Printf("已合并导出: %s\n", outputPath)
	return nil
}

func GetSheetInfo(filePath string) ([]SheetInfo, error) {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, err
	}
	defer f.Close()

	sheets := f.GetSheetList()
	sheetInfos := make([]SheetInfo, 0, len(sheets))

	for _, sheet := range sheets {
		rows, _ := f.GetRows(sheet)
		maxCols := 0
		for _, row := range rows {
			if len(row) > maxCols {
				maxCols = len(row)
			}
		}

		info := SheetInfo{
			Name:     sheet,
			RowCount: len(rows),
			ColCount: maxCols,
		}
		sheetInfos = append(sheetInfos, info)
		fmt.Printf("Sheet: %s, 行数: %d, 列数: %d\n", info.Name, info.RowCount, info.ColCount)
	}

	return sheetInfos, nil
}

func QuickSplit(inputFile, outputDir string) error {
	config := SplitConfig{
		InputFile:    inputFile,
		OutputDir:    outputDir,
		CopyStyles:   true,
		CopyFormulas: false,
	}
	return RunSplit(config)
}

func QuickMerge(inputFile, outputDir string) error {
	config := SplitConfig{
		InputFile:  inputFile,
		OutputDir:  outputDir,
		CopyStyles: true,
	}
	return RunMerge(config)
}

func RunSplit(config SplitConfig) error {
	fmt.Printf("开始拆分 Excel 文件: %s\n", config.InputFile)
	fmt.Println("\nSheet 信息:")
	if _, err := GetSheetInfo(config.InputFile); err != nil {
		fmt.Printf("获取 Sheet 信息失败: %v\n", err)
	}

	fmt.Println("\n开始导出...")
	return SplitExcelSheetsAdvanced(config)
}

func RunMerge(config SplitConfig) error {
	fmt.Printf("开始按表头合并 Excel 文件: %s\n", config.InputFile)
	fmt.Println("\nSheet 信息:")
	if _, err := GetSheetInfo(config.InputFile); err != nil {
		fmt.Printf("获取 Sheet 信息失败: %v\n", err)
	}

	fmt.Println("\n开始合并导出...")
	return MergeExcelSheetsByHeader(config)
}

func getSheetsToProcess(srcFile *excelize.File, sheets []string) ([]string, error) {
	if len(sheets) > 0 {
		return sheets, nil
	}

	allSheets := srcFile.GetSheetList()
	if len(allSheets) == 0 {
		return nil, fmt.Errorf("文件中没有找到任何 Sheet")
	}

	return allSheets, nil
}

func extractHeader(rows [][]string) (int, []string, bool) {
	for rowIdx, row := range rows {
		header := normalizeRow(row)
		if len(header) == 0 {
			continue
		}
		return rowIdx, header, true
	}

	return 0, nil, false
}

func normalizeRow(row []string) []string {
	normalized := make([]string, len(row))
	for i, cell := range row {
		normalized[i] = strings.TrimSpace(cell)
	}

	last := len(normalized)
	for last > 0 && normalized[last-1] == "" {
		last--
	}

	return normalized[:last]
}

func headerSignature(header []string) string {
	return strings.Join(normalizeRow(header), "\x1f")
}

func writeRow(dstFile *excelize.File, sheetName string, rowNumber int, row []string) error {
	for colIdx, cellValue := range row {
		cell, err := excelCellName(rowNumber, colIdx+1)
		if err != nil {
			return err
		}

		if err := dstFile.SetCellValue(sheetName, cell, cellValue); err != nil {
			return fmt.Errorf("写入单元格 %s 失败: %w", cell, err)
		}
	}

	return nil
}

func copyColumnWidths(srcFile, dstFile *excelize.File, srcSheet, dstSheet string) {
	cols, err := srcFile.GetCols(srcSheet)
	if err != nil {
		return
	}

	for i := range cols {
		colLetter, err := excelize.ColumnNumberToName(i + 1)
		if err != nil {
			continue
		}

		width, err := srcFile.GetColWidth(srcSheet, colLetter)
		if err == nil && width > 0 {
			_ = dstFile.SetColWidth(dstSheet, colLetter, colLetter, width)
		}
	}
}

func copyRowHeights(srcFile, dstFile *excelize.File, srcSheet, dstSheet string, rowCount int) {
	for rowIdx := 1; rowIdx <= rowCount; rowIdx++ {
		height, err := srcFile.GetRowHeight(srcSheet, rowIdx)
		if err == nil && height > 0 {
			_ = dstFile.SetRowHeight(dstSheet, rowIdx, height)
		}
	}
}

func excelCellName(rowNumber, colNumber int) (string, error) {
	colLetter, err := excelize.ColumnNumberToName(colNumber)
	if err != nil {
		return "", fmt.Errorf("转换列号失败: %w", err)
	}

	return fmt.Sprintf("%s%d", colLetter, rowNumber), nil
}

func generateSplitFileName(baseName, sheetName string, addTimestamp bool) string {
	cleanSheetName := sanitizeSheetName(sheetName)
	if addTimestamp {
		timestamp := time.Now().Format("20060102_150405")
		return fmt.Sprintf("%s_%s_%s.xlsx", baseName, cleanSheetName, timestamp)
	}

	return fmt.Sprintf("%s_%s.xlsx", baseName, cleanSheetName)
}

func generateMergedFileName(sheets []sheetData, addTimestamp bool) string {
	parts := make([]string, 0, len(sheets))
	for _, sheet := range sheets {
		parts = append(parts, sanitizeSheetName(sheet.Name))
	}

	name := strings.Join(parts, "_")
	if name == "" {
		name = "merged"
	}

	if addTimestamp {
		timestamp := time.Now().Format("20060102_150405")
		return fmt.Sprintf("%s_%s.xlsx", name, timestamp)
	}

	return fmt.Sprintf("%s.xlsx", name)
}

func sanitizeSheetName(name string) string {
	invalidChars := []string{"/", "\\", ":", "*", "?", "\"", "<", ">", "|"}
	result := strings.TrimSpace(name)
	for _, char := range invalidChars {
		result = strings.ReplaceAll(result, char, "_")
	}

	if result == "" {
		return "sheet"
	}

	return result
}

func parseSheets(raw string) []string {
	if strings.TrimSpace(raw) == "" {
		return nil
	}

	parts := strings.Split(raw, ",")
	sheets := make([]string, 0, len(parts))
	for _, part := range parts {
		name := strings.TrimSpace(part)
		if name != "" {
			sheets = append(sheets, name)
		}
	}

	return sheets
}

func main() {
	inputFile := flag.String("input", "input.xlsx", "待处理的 Excel 文件路径")
	outputDir := flag.String("output", "./output", "导出目录")
	mode := flag.String("mode", "merge", "导出模式: split 或 merge")
	sheets := flag.String("sheets", "", "只处理指定 Sheet，多个用英文逗号分隔")
	copyStyles := flag.Bool("copy-styles", true, "是否复制基础样式")
	copyFormulas := flag.Bool("copy-formulas", false, "split 模式下是否保留公式")
	addTimestamp := flag.Bool("timestamp", false, "是否在输出文件名中追加时间戳")
	flag.Parse()

	config := SplitConfig{
		InputFile:      *inputFile,
		OutputDir:      *outputDir,
		Sheets:         parseSheets(*sheets),
		CopyStyles:     *copyStyles,
		CopyFormulas:   *copyFormulas,
		AddTimestamp:   *addTimestamp,
		PreserveFormat: false,
	}

	var err error
	switch strings.ToLower(strings.TrimSpace(*mode)) {
	case "split":
		err = RunSplit(config)
	case "merge":
		err = RunMerge(config)
	default:
		err = fmt.Errorf("不支持的模式 %q，请使用 split 或 merge", *mode)
	}

	if err != nil {
		fmt.Printf("错误: %v\n", err)
		os.Exit(1)
	}
}
