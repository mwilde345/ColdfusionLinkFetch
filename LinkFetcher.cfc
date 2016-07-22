<cfscript>
function getHyperlinks(spreadsheet){
	var linksArray = arraynew(1);
	var mainFile = createObject("java","java.io.File").init(spreadsheet);
	var inputStream = createObject("java","java.io.FileInputStream").init(mainFile);
	var workbook = createObject("java","org.apache.poi.xssf.usermodel.XSSFWorkbook").init(inputStream); //change to hssf for .xls
	var poiSpreadsheet = createObject("java","org.apache.poi.xssf.usermodel.XSSFSheet"); //hssf for .xls instead of xssf for .xlsx
	var poiSpreadsheet = workbook.getSheetAt(0);
	var rowIterator = createObject("java","java.util.Iterator");
	rowIterator = poiSpreadsheet.iterator();
	while(rowIterator.hasNext()){
		var row = createObject("java","org.apache.poi.ss.usermodel.Row");
		row = rowIterator.next();
		var cellTraverse = createObject("java","java.util.Iterator");
		cellTraverse = row.cellIterator();
		var colCounter = 1;
		var rowStruct = structNew();
		while(cellTraverse.hasNext()){
			var cell = createObject("java","org.apache.poi.ss.usermodel.Cell");
			cell = cellTraverse.next();
			var df = createObject("java","org.apache.poi.ss.usermodel.DataFormatter").init();
			var theLink = createObject("java","org.apache.poi.common.usermodel.Hyperlink
			theLink = cell.getHyperlink();
			if(isDefined("theLink")){
				rowStruct["col_"&colCounter] = theLink.getAddress();
			}
			colCounter++;
		}
		arrayappend(linksarray,rowStruct);
	}
	return linksArray;
}
</cfscript>
