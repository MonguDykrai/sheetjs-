import XLSX from 'xlsx';

type ExportToExcelOptions = {
  columnWidth: number[]; // 列宽
  filename: string; // 文件名
  header: string[]; // table header 表头
  rows: object[]; // table row data 行数据
  extension?: '.xlsx'; // 文件扩展名
};

const isParametersValid = (
  filename: string,
  header: string[],
  rows: object[]
) => {
  try {
    if (
      !(
        Array.isArray(header) &&
        Array.isArray(rows) &&
        typeof filename === 'string'
      )
    )
      throw new Error(
        '参数不合法: filename(string), header(string[]), rows(object[])'
      );
  } catch (error) {
    if (error instanceof Error) return new Error(error.message);
  }
};

/**
 * 导出 Excel 文件
 *
 * 执行过程中会对各项参数进行校验，校验失败则抛错，成功则下载文件。
 * @param ExportToExcelOptions filename, header, rows, extension(optional)
 * @returns
 */
export default function exportToExcel({
  columnWidth,
  filename = '',
  header = [],
  rows = [],
  extension = '.xlsx',
}: ExportToExcelOptions) {
  try {
    const validationReport = isParametersValid(filename, header, rows);
    if (validationReport instanceof Error)
      throw new Error(validationReport.message);

    const $filename = filename + extension;
    const $worksheet_name = 'Sheet1'; // 工作表名

    if (rows.length === 0) return; // 无数据则不做任何处理

    const data = rows.map((row) => Object.values(row)); // [{name:'李雷',id:'1'},{name:'韩梅梅',id:'2'}] => [['李雷','1'],['韩梅梅','2']]
    data.unshift(header); // [['姓名','学号'],['李雷','1'],['韩梅梅','2']]

    let wb = XLSX.utils.book_new(); // workbook 工作簿对象
    let ws = XLSX.utils.aoa_to_sheet(data); // worksheet 工作表对象
    if (columnWidth.length > 0)
      ws['!cols'] = columnWidth.map((wch) => ({ wch }));
    XLSX.utils.book_append_sheet(wb, ws, $worksheet_name); // Append a worksheet to a workbook

    XLSX.writeFile(wb, $filename); // 写入并下载 Excel 文件
  } catch (error) {
    if (error instanceof Error) throw new Error(error.message);
  }
}
