#ifndef SQLRESULTTABLE_H
#define SQLRESULTTABLE_H

#include <wx/grid.h>
#include "wxPostgreSQL.h"
#include "wf_stdafx.h"


// we cannot derive from wxGridCellAttr because destructor is private but not virtual
class sqlCellAttr
{
public:
	sqlCellAttr()
	{
		attr = new wxGridCellAttr;

	}
	~sqlCellAttr()
	{
		attr->DecRef();
	}

	wxGridCellAttr *attr;

};

class sqlResultTable;
class sqlResultGrid : public wxGrid
{
public:
    sqlResultGrid(wxWindow *parent, wxWindowID id, const wxPoint &pos, const wxSize &size);
    sqlResultGrid();


    wxString GetExportLine(int row);
	wxString GetExportLine(int row, wxArrayInt cols);
	wxString GetExportLine(int row, int col1, int col2);

    sqlResultTable *GetTable()
    {
        return (sqlResultTable *)wxGrid::GetTable();
    }
    //wxSize GetBestSize(int row, int col);
    void ResizeEditor(int row, int col);
    wxArrayInt GetSelectedRows() const;
    void show_tooltip( wxMouseEvent& event );
    int Copy();
 //   DECLARE_DYNAMIC_CLASS(sqlResultGrid)
	DECLARE_EVENT_TABLE()

private:
    void OnMouseWheel(wxMouseEvent &event);
    void OnCopy(wxCommandEvent &event);

};

class sqlResultTable : public wxGridTableBase
{
public:
    sqlResultTable(wxPostgreSQLresult * _res, wxString &tablename, wxString &header,  bool isEditable);
    virtual ~sqlResultTable();

    void storeling(wxArrayString & line);
    int GetNumberRows() { return rows_num; }
    int GetNumberCols() { return cols_num; }
    wxString GetValue( int row, int col );
    void SetValue( int row, int col, const wxString& s );

    void storedata(int irow, int icol);

    wxString GetColLabelValue (int col);

    void SetColLabelValue(int col, const wxString &label);
    void SetRowLabelValue(int row, const wxString &label);
    wxString GetRowLabelValue(int row);
    bool AppendRows (size_t numRows=1);
    bool DeleteRows( size_t pos = 0, size_t numRows = 1 );
    bool InsertRows( size_t pos = 0, size_t numRows = 1 );

    wxGridCellAttr *GetAttr(int row, int col, wxGridCellAttr::wxAttrKind  kind);

    void SetLabelValue(wxString label);

    wxArrayString  GetPrimaryKeyValue(int row);

    void IsCanSave(bool bsave)
    {
        b_Save = bsave;
    }

protected:
private:
    sqlCellAttr *row_attr;
    bool b_Save;
    wxPostgreSQLresult* res;
    wxString table_name;
    wxArrayString primay_keys;
    wxArrayInt colTypes;
    wxArrayString colNames;

    wxGridStringArray m_data;
    wxArrayString col_Labels, row_Labels;

    int cols_num;
    long rows_num;
};

#endif // SQLRESULTTABLE_H
