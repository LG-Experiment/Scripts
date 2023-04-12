function main(workbook: ExcelScript.Workbook) {
	let tasks = workbook.getTable("Tasks");
	// Apply checked items filter on table tasks column Status
	tasks.getColumnByName("Status").getFilter().applyValuesFilter(["Done"]);
	let archive = workbook.getWorksheet("Archive");
	let get_visible_view = tasks.getRangeBetweenHeaderAndTotal().getVisibleView();
	if (get_visible_view.getRowCount() > 0) {
		let paste_cell = archive.getRange("A1").getRangeEdge(ExcelScript.KeyboardDirection.down).getOffsetRange(1, 0);
		// paste done tasks
		let paste_location = paste_cell.getResizedRange(get_visible_view.getRowCount()-1,get_visible_view.getColumnCount()-1);
		paste_location.setValues(get_visible_view.getValues());
	}
	// Clear filter on table tasks column "Status"
	tasks.getColumnByName("Status").getFilter().clear();
	// delete done tasks
	let tasks_rows = tasks.getRowCount();
	// console.log (tasks_rows)
	let tasks_range = tasks.getRangeBetweenHeaderAndTotal().getValues();
	for (var i=tasks_rows-1; i>-1; i--) {
		if (tasks_range[i][5].toString().toUpperCase() == "DONE") {
			tasks.deleteRowsAt(i,1);
		}
	}
}
