package main

import (
	"flag"
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

// SplitConfig 定义了程序执行拆分或合并时会使用到的全部配置项。
// 这个结构体会在 main 中组装完成，然后传递给后续的业务方法。
type SplitConfig struct {
	InputFile      string
	OutputDir      string
	Sheets         []string
	CopyStyles     bool
	CopyFormulas   bool
	AddTimestamp   bool
	PreserveFormat bool
}

// SheetInfo 表示一个 Sheet 的基础信息。
// 目前主要用于在程序开始处理前打印概览，方便确认输入文件内容是否正确。
type SheetInfo struct {
	Name     string
	RowCount int
	ColCount int
}

// sheetData 保存单个 Sheet 在“合并模式”下需要用到的内容。
// 它会记录 Sheet 名称、识别出的表头，以及真正要参与合并的数据行。
type sheetData struct {
	Name     string
	Header   []string
	DataRows [][]string
}

// mergeGroup 表示一组“表头完全一致”的 Sheet。
// 合并模式会先按表头把多个 Sheet 分组，再把同一组导出为一个 Excel 文件。
type mergeGroup struct {
	Header []string
	Sheets []sheetData
}

// SplitExcelSheets 是一个便捷入口。
// 它使用默认配置执行“按 Sheet 拆分导出”，适合不需要自定义参数的场景。
func SplitExcelSheets(inputFile, outputDir string) error {
	return SplitExcelSheetsAdvanced(SplitConfig{
		InputFile:    inputFile,
		OutputDir:    outputDir,
		CopyStyles:   true,
		CopyFormulas: false,
	})
}

// SplitExcelSheetsAdvanced 根据传入配置执行拆分逻辑。
// 这个方法会打开源 Excel、确定要处理的 Sheet 列表，并逐个导出为独立文件。
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

// MergeExcelSheetsByHeader 根据表头内容对多个 Sheet 进行分组并导出。
// 如果两个或多个 Sheet 的表头一致，就会被合并到同一个结果文件中。
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

// exportSingleSheet 负责把一个指定的 Sheet 写入到新的 Excel 文件中。
// 这是拆分模式最核心的方法，会复制单元格内容，并按配置决定是否复制公式和样式。
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

// exportMergedGroup 把同一组表头一致的 Sheet 合并写入到一个新文件。
// 它只写一次表头，随后按顺序追加每个 Sheet 的数据行，并生成拼接后的文件名。
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

// GetSheetInfo 读取 Excel 中每个 Sheet 的基础统计信息。
// 这个方法主要用于输出日志，帮助用户在正式导出前先看到文件结构概览。
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

// QuickSplit 提供一个使用默认拆分参数的快捷调用方式。
// 当外部只关心输入和输出目录，而不关心其他细节时，可以直接调用它。
func QuickSplit(inputFile, outputDir string) error {
	config := SplitConfig{
		InputFile:    inputFile,
		OutputDir:    outputDir,
		CopyStyles:   true,
		CopyFormulas: false,
	}
	return RunSplit(config)
}

// QuickMerge 提供一个使用默认合并参数的快捷调用方式。
// 它会构造基础配置，并交给 RunMerge 执行完整流程。
func QuickMerge(inputFile, outputDir string) error {
	config := SplitConfig{
		InputFile:  inputFile,
		OutputDir:  outputDir,
		CopyStyles: true,
	}
	return RunMerge(config)
}

// RunSplit 是拆分模式的流程入口。
// 它会先打印输入文件的 Sheet 概览，再调用真正的拆分导出方法。
func RunSplit(config SplitConfig) error {
	fmt.Printf("开始拆分 Excel 文件: %s\n", config.InputFile)
	fmt.Println("\nSheet 信息:")
	if _, err := GetSheetInfo(config.InputFile); err != nil {
		fmt.Printf("获取 Sheet 信息失败: %v\n", err)
	}

	fmt.Println("\n开始导出...")
	return SplitExcelSheetsAdvanced(config)
}

// RunMerge 是合并模式的流程入口。
// 它和 RunSplit 类似，先打印文件概览，再执行按表头分组的合并导出逻辑。
func RunMerge(config SplitConfig) error {
	fmt.Printf("开始按表头合并 Excel 文件: %s\n", config.InputFile)
	fmt.Println("\nSheet 信息:")
	if _, err := GetSheetInfo(config.InputFile); err != nil {
		fmt.Printf("获取 Sheet 信息失败: %v\n", err)
	}

	fmt.Println("\n开始合并导出...")
	return MergeExcelSheetsByHeader(config)
}

// getSheetsToProcess 决定本次任务到底处理哪些 Sheet。
// 如果用户显式传入了 Sheet 名列表，就只处理这些；否则处理文件中的全部 Sheet。
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

// extractHeader 从一个 Sheet 的全部行数据中识别表头。
// 当前规则是找到第一行“非空行”作为表头，并返回该表头所在行号和内容。
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

// normalizeRow 对一整行数据进行标准化处理。
// 它会去掉单元格首尾空格，并裁掉行尾连续的空列，便于后续比较表头是否一致。
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

// headerSignature 把一个表头切片转换成可比较的签名字符串。
// 合并模式会用这个签名作为 map key，从而快速判断两个 Sheet 的表头是否一致。
func headerSignature(header []string) string {
	return strings.Join(normalizeRow(header), "\x1f")
}

// writeRow 把一整行数据按列顺序写入目标 Excel 的指定行。
// 这是导出时的通用写行工具，拆分和合并流程都会复用它的底层单元格定位逻辑。
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

// copyColumnWidths 把源 Sheet 的列宽复制到目标 Sheet。
// 这样导出的文件在视觉上会更接近原始 Excel，减少列宽错乱的问题。
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

// copyRowHeights 把源 Sheet 的行高复制到目标 Sheet。
// 目前拆分模式会用它保留基础展示效果，避免导出后行高全部回到默认值。
func copyRowHeights(srcFile, dstFile *excelize.File, srcSheet, dstSheet string, rowCount int) {
	for rowIdx := 1; rowIdx <= rowCount; rowIdx++ {
		height, err := srcFile.GetRowHeight(srcSheet, rowIdx)
		if err == nil && height > 0 {
			_ = dstFile.SetRowHeight(dstSheet, rowIdx, height)
		}
	}
}

// excelCellName 根据行号和列号生成 Excel 单元格坐标。
// 例如传入第 2 行第 3 列时，会返回 C2。
func excelCellName(rowNumber, colNumber int) (string, error) {
	colLetter, err := excelize.ColumnNumberToName(colNumber)
	if err != nil {
		return "", fmt.Errorf("转换列号失败: %w", err)
	}

	return fmt.Sprintf("%s%d", colLetter, rowNumber), nil
}

// generateSplitFileName 为拆分模式生成输出文件名。
// 文件名由原始文件名、Sheet 名以及可选时间戳组成。
func generateSplitFileName(baseName, sheetName string, addTimestamp bool) string {
	cleanSheetName := sanitizeSheetName(sheetName)
	if addTimestamp {
		timestamp := time.Now().Format("20060102_150405")
		return fmt.Sprintf("%s_%s_%s.xlsx", baseName, cleanSheetName, timestamp)
	}

	return fmt.Sprintf("%s_%s.xlsx", baseName, cleanSheetName)
}

// generateMergedFileName 为合并模式生成输出文件名。
// 它会把同组中所有 Sheet 名拼接起来，让结果文件能直接反映来源。
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

// sanitizeSheetName 清理文件名中不允许出现的字符。
// 因为 Sheet 名会直接参与输出文件名生成，所以这里需要先做一次安全转换。
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

// parseSheets 把命令行传入的 Sheet 字符串解析成切片。
// 用户可以用英文逗号传多个 Sheet 名，这里会负责切分和去除空白内容。
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

// resolveInputFile 决定最终要处理的输入文件。
// 如果用户传了 -input 就直接使用；如果没传，就自动扫描当前目录并尝试找到唯一的 Excel 文件。
func resolveInputFile(inputFile string) (string, error) {
	inputFile = strings.TrimSpace(inputFile)
	if inputFile != "" {
		return inputFile, nil
	}

	entries, err := os.ReadDir(".")
	if err != nil {
		return "", fmt.Errorf("读取当前目录失败: %w", err)
	}

	matches := make([]string, 0)
	for _, entry := range entries {
		if entry.IsDir() {
			continue
		}

		name := entry.Name()
		if strings.HasPrefix(name, "~$") {
			continue
		}

		switch strings.ToLower(filepath.Ext(name)) {
		case ".xlsx", ".xlsm", ".xltx", ".xltm":
			matches = append(matches, name)
		}
	}

	sort.Strings(matches)

	switch len(matches) {
	case 0:
		return "", fmt.Errorf("当前目录没有找到 Excel 文件，请使用 -input 指定文件")
	case 1:
		fmt.Printf("未传入 inputFile，自动使用当前目录中的 Excel 文件: %s\n", matches[0])
		return matches[0], nil
	default:
		return "", fmt.Errorf("当前目录找到多个 Excel 文件: %s，请使用 -input 指定要处理的文件", strings.Join(matches, ", "))
	}
}

// main 是程序的命令行入口。
// 它负责解析参数、补齐最终配置、选择运行模式，并把错误打印到终端。
func main() {
	inputFile := flag.String("input", "", "待处理的 Excel 文件路径；不传时自动扫描当前目录")
	outputDir := flag.String("output", "./output", "导出目录")
	mode := flag.String("mode", "merge", "导出模式: split 或 merge")
	sheets := flag.String("sheets", "", "只处理指定 Sheet，多个用英文逗号分隔")
	copyStyles := flag.Bool("copy-styles", true, "是否复制基础样式")
	copyFormulas := flag.Bool("copy-formulas", false, "split 模式下是否保留公式")
	addTimestamp := flag.Bool("timestamp", false, "是否在输出文件名中追加时间戳")
	flag.Parse()

	resolvedInputFile, err := resolveInputFile(*inputFile)
	if err != nil {
		fmt.Printf("错误: %v\n", err)
		os.Exit(1)
	}

	config := SplitConfig{
		InputFile:      resolvedInputFile,
		OutputDir:      *outputDir,
		Sheets:         parseSheets(*sheets),
		CopyStyles:     *copyStyles,
		CopyFormulas:   *copyFormulas,
		AddTimestamp:   *addTimestamp,
		PreserveFormat: false,
	}

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
