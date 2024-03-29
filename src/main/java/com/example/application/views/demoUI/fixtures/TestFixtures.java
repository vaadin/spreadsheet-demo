package com.example.application.views.demoUI.fixtures;

/**
 * Test fixtures for server-side Spreadsheet manipulation
 *
 */
public enum TestFixtures {
    FirstColumnWidth(FirstColumnWidthFixture.class), 
	PopupButton(PopupButtonFixture.class),
	TabsheetPopupButton(TabsheetPopupButtonFixture.class),
	TablePopupButton(TablePopupButtonFixture.class),
	SpreadsheetTable(SpreadsheetTableFixture.class),
	Comments(CommentFixture.class), 
	AddOrRemoveComment(AddOrRemoveCommentFixture.class),
	Formats(FormatsFixture.class), 
	StyleMergeReigions(StyleMergeReigions.class),
	RemoveFixture(RemoveFixture.class),
	DefaultStyleUnlocked(DefaultStyleUnlockedFixture.class),
	HideSecondRow(HideSecondRowFixture.class),
	LargeSpreadsheet(LargeSpreadsheetFixture.class),
	ColumnToggle(ColumnToggleFixture.class),
	RowToggle(RowToggleFixture.class),
	DeletionHandler(DeletionHandlerFixture.class),
	Selection(SelectionFixture.class),
	MergeCells(CellMergeFixture.class),
	ValueChangeHandler(ValueHandlerFixture.class),
	Rename(RenameFixture.class),
	CreateSheet(SheetsFixture.class),
	Styles(StylesFixture.class),
	LockCell(LockCellFixture.class),
	Action(ActionFixture.class),
	InsertRow(InsertRowFixture.class),
    DeleteRow(DeleteRowFixture.class),
    RowHeaderDoubleClick(RowHeaderDoubleClickFixture.class);

    public final SpreadsheetFixtureFactory factory;

    TestFixtures(SpreadsheetFixtureFactory factory) {
        this.factory = factory;
    }

    TestFixtures(Class<? extends SpreadsheetFixture> fixtureClass) {
        this(new ClassFixtureFactory(fixtureClass));
    }
}
