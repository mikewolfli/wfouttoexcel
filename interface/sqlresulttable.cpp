#include "sqlresulttable.h"
#include "wfApp.h"
#include <wx/tipwin.h>
#include <wx/clipbrd.h>


BEGIN_EVENT_TABLE(sqlResultGrid, wxGrid)
//	EVT_MOUSEWHEEL(sqlResultGrid::OnMouseWheel)
//	EVT_MOTION(sqlResultGrid::show_tooltip)
END_EVENT_TABLE()

//IMPLEMENT_DYNAMIC_CLASS(sqlResultGrid, wxGrid)

sqlResultGrid::sqlResultGrid()
{

}

sqlResultGrid::sqlResultGrid(wxWindow *parent, wxWindowID id, const wxPoint &pos, const wxSize &size)
    : wxGrid(parent, id, pos, size, wxWANTS_CHARS | wxVSCROLL | wxHSCROLL, _T(""))
{
    GetGridWindow()->Connect(GetGridWindow()->GetId(),wxEVT_MOTION,wxMouseEventHandler(sqlResultGrid::show_tooltip),NULL,this);
    GetGridWindow()->Connect(GetGridWindow()->GetId(), wxEVT_MOUSEWHEEL, wxMouseEventHandler(sqlResultGrid::OnMouseWheel),NULL, this);
}


void sqlResultGrid::OnCopy(wxCommandEvent &ev)
{
	Copy();
}

int sqlResultGrid::Copy()
{
	wxString str;
	int copied = 0;
	size_t i;



	if (GetSelectedRows().GetCount())
	{

		wxArrayInt rows = GetSelectedRows();

		for (i = 0 ; i < rows.GetCount() ; i++)
		{
			str.Append(GetExportLine(rows.Item(i)));

			if (rows.GetCount() > 1)
				str.Append(END_OF_LINE);
		}

		copied = rows.GetCount();
	}
	else if (GetSelectedCols().GetCount())
	{
		wxArrayInt cols = GetSelectedCols();
		size_t numRows = GetNumberRows();

     	for (i = 0 ; i < numRows ; i++)
		{
			str.Append(GetExportLine(i, cols));

			if (numRows > 1)
				str.Append(END_OF_LINE);
		}

		copied = numRows;
	}
	else if (GetSelectionBlockTopLeft().GetCount() > 0 &&
	         GetSelectionBlockBottomRight().GetCount() > 0)
	{
		unsigned int x1, x2, y1, y2;

		x1 = GetSelectionBlockTopLeft()[0].GetCol();
		x2 = GetSelectionBlockBottomRight()[0].GetCol();
		y1 = GetSelectionBlockTopLeft()[0].GetRow();
		y2 = GetSelectionBlockBottomRight()[0].GetRow();

		for (i = y1; i <= y2; i++)
		{
			str.Append(GetExportLine(i, x1, x2));

			if (y2 > y1)
				str.Append(END_OF_LINE);
		}

		copied = y2 - y1 + 1;
	}
	else
	{
		int row, col;

		row = GetGridCursorRow();
		col = GetGridCursorCol();

		str.Append(GetExportLine(row, col, col));
		copied = 1;
	}

	if (copied && wxTheClipboard->Open())
	{
		wxTheClipboard->SetData(new wxTextDataObject(str));
		wxTheClipboard->Close();
	}
	else
	{
		copied = 0;
	}

	return copied;
}

wxString sqlResultGrid::GetExportLine(int row)
{
	return GetExportLine(row, 0, GetNumberCols() - 1);
}


wxString sqlResultGrid::GetExportLine(int row, int col1, int col2)
{
	wxArrayInt cols;
	wxString str;
	int i;

	if (col2 < col1)
		return str;

	cols.Alloc(col2 - col1 + 1);
	for (i = col1; i <= col2; i++)
	{
		cols.Add(i);
	}

	return GetExportLine(row, cols);
}

wxString sqlResultGrid::GetExportLine(int row, wxArrayInt cols)
{
	wxString str;
	unsigned int col;

	if (GetNumberCols() == 0)
		return str;

	for (col = 0 ; col < cols.Count() ; col++)
	{
		if (col > 0)
//			str.Append(settings->GetCopyColSeparator());
            str.Append(wxT("\t"));

		wxString text = GetCellValue(row, cols[col]);

		bool needQuote  = false;

		str.Append(text);

	}
	return str;
}

void sqlResultGrid::show_tooltip( wxMouseEvent& event )
{
  // get cell coords
  wxPoint point = ScreenToClient(wxGetMousePosition());
//  wxPoint point = ScreenToClient(wxPoint(event.GetX(),event.GetY()));
  int x, y;
  CalcUnscrolledPosition(point.x,point.y,&x,&y);


  int row = YToRow(y-GetColLabelSize());
  int col = XToCol(x-GetRowLabelSize());

  if( col == wxNOT_FOUND || row == wxNOT_FOUND )
    {
      event.Skip();
      return;
    }

  // fetch cell value and type
  wxString val = GetCellValue(row, col);

  // show tooltip
  GetGridWindow()->SetToolTip(val);

  event.Skip();
}

wxArrayInt sqlResultGrid::GetSelectedRows() const
{
    wxArrayInt rows, rows2;

    wxGridCellCoordsArray tl = GetSelectionBlockTopLeft(), br = GetSelectionBlockBottomRight();

    int maxCol = ((sqlResultGrid *)this)->GetNumberCols() - 1;
    size_t i;
    for (i = 0 ; i < tl.GetCount() ; i++)
    {
        wxGridCellCoords c1 = tl.Item(i), c2 = br.Item(i);
        if (c1.GetCol() != 0 || c2.GetCol() != maxCol)
            return rows2;

        int j;
        for (j = c1.GetRow() ; j <= c2.GetRow() ; j++)
            rows.Add(j);
    }

    rows2 = wxGrid::GetSelectedRows();

    rows.Sort(ArrayCmp);
    rows2.Sort(ArrayCmp);

    size_t i2 = 0, cellRowMax = rows.GetCount();

    for (i = 0 ; i < rows2.GetCount() ; i++)
    {
        int row = rows2.Item(i);
        while (i2 < cellRowMax && rows.Item(i2) < row)
            i2++;
        if (i2 == cellRowMax || row != rows.Item(i2))
            rows.Add(row);
    }

    return rows;
}

void sqlResultGrid::OnMouseWheel(wxMouseEvent &event)
{
	if (event.ControlDown())
	{
		wxFont fontlabel = GetLabelFont();
		wxFont fontcells = GetDefaultCellFont();
		if (event.GetWheelRotation() > 0)
		{
			fontlabel.SetPointSize(fontlabel.GetPointSize() - 1);
			fontcells.SetPointSize(fontcells.GetPointSize() - 1);
		}
		else
		{
			fontlabel.SetPointSize(fontlabel.GetPointSize() + 1);
			fontcells.SetPointSize(fontcells.GetPointSize() + 1);
		}
		SetLabelFont(fontlabel);
		SetDefaultCellFont(fontcells);
		SetColLabelSize(fontlabel.GetPointSize() * 4);
		SetDefaultRowSize(fontcells.GetPointSize() * 2);
		for (int index = 0; index < GetNumberCols(); index++)
			SetColSize(index, -1);
		ForceRefresh();
	}
	else
		event.Skip();
}


sqlResultTable::sqlResultTable(wxPostgreSQLresult * _res, wxString &tablename, wxString &header, bool isEditable)
{
    //ctor
    res=_res;
    table_name = tablename;

    rows_num = res->GetRowsNumber();
    cols_num = res->GetNumberFields();

    SetLabelValue(header);
    b_Save = false;

    int i,j;

    for (i=0; i<cols_num; i++)
    {
        colTypes.Add(_res->ColTypeOID(i));
        colNames.Add(_res->ColName(i));
    }

//    m_data.Alloc(rows_num);

    wxArrayString line;

    res->MoveFirst();

    wxString str;


    row_attr = new sqlCellAttr[rows_num];
    bool b_urgent=false;
    bool b_freeze = false;
//    wxString str_prj_catalog;

    for(i=0; i<rows_num; i++)
    {
        str= wxString::Format("%ld",i+1);
        row_Labels.Add(str);
        for(j=0; j<cols_num; j++)
        {
            switch(colTypes.Item(j))
            {
            case PGOID_TYPE_BOOL:
                if(res->GetBool(j))
                    str=wxT("Y");
                else str.Empty();

                if(colNames.Item(j)== wxT("is_urgent"))
                {
                    if(res->GetBool(j))
                        b_urgent = true;
                    else b_urgent = false;
                }
                break;
            case PGOID_TYPE_INT8:
            case PGOID_TYPE_SERIAL8:
            case PGOID_TYPE_INT2:
            case PGOID_TYPE_INT4:
            case PGOID_TYPE_SERIAL:
                if(colNames.Item(j)== wxT("status"))
                {
                    _status t_status = (_status)res->GetInt(j);
                    if(t_status == IS_FREEZED)
                        b_freeze = true;
                    else b_freeze = false;
                    switch(t_status)
                    {
                    case IS_CREATED:
                        str=wxT("CRTD");
                        break;
                    case IS_DELETED:
                        str=wxT("DEL");
                        break;
                    case IS_ACTIVE:
                        str=wxT("ACTIVE");
                        break;
                    case IS_FINISHED:
                        str= wxT("FININISH");
                        break;
                    case IS_FREEZED:
                        str=wxT("FREEZED");
                        break;
                    case IS_CLOSED:
                        str= wxT("CLOSED");
                        break;
                    case IS_IN_DOC:
                        str= wxT("HISTORY");
                    default:
                        break;
                    }
                }
                else if(colNames.Item(j)== wxT("project_catalog"))
                {
                    prj_type _pjtype = (prj_type)res->GetInt(j);
                    switch(_pjtype)
                    {
                    case COMMON_PROJECT:
                        str = wxT("Common Project");
                        break;
                    case HIGH_SPEED_PROJECT:
                        str = wxT("High-Speed Project");
                        break;
                    case SPECIAL_PROJECT:
                        str = wxT("Special Project");
                        break;
                    case MAJOR_PROJECT:
                        str = wxT("Major Project");
                        break;
                    case PRE_ENGINEERING:
                        str = wxT("Pre_engineering");
                        break;
                    case JY:
                        str = wxT("JY");
                    default:
                        break;
                    }

 //                   str_prj_catalog = str;
                }
                else if(colNames.Item(j)== wxT("nonstd_level"))
                {
                    non_std_level _non_std_level = (non_std_level)res->GetInt(j);
                    switch(_non_std_level)
                    {
                case complete_standard:
                           str = wxT("Complete Standard");
                           break;
                case option_standard:
                           str = wxT("Option Standard");
                           break;
                case standard_configuration_without_pc:
                           str = wxT("Standard Configuration Without PC");
                           break;
                case  simple_non_standard:
                           str = wxT("Simple Non-Standard");
                           break;
                case complex_non_standard:
                           str = wxT("Complex Non-Standard");
                           break;
                default:
                    break;
                    }
                }
                else str = res->GetVal(j);

                break;
            case PGOID_TYPE_DATE:
            case PGOID_TYPE_TIME:
            case PGOID_TYPE_TIMETZ:
            case PGOID_TYPE_TIMESTAMP:
            case PGOID_TYPE_TIMESTAMPTZ:
            case PGOID_TYPE_INTERVAL:
                if(!res->GetVal(j).IsEmpty())
                    str = res->GetDate(j).Format("%Y-%m-%d");
                else str = wxEmptyString;
                break;
            default:
                str = res->GetVal(j);
                break;
            }
            line.Add(str);
        }

        if(!isEditable)
            row_attr[i].attr->SetReadOnly();

        if(b_freeze)
            row_attr[i].attr->SetBackgroundColour(*wxRED);
        else if(b_urgent)
            row_attr[i].attr->SetBackgroundColour(*wxGREEN);
        else row_attr[i].attr->SetBackgroundColour(*wxWHITE);


        m_data.Add(line);
        line.Clear();
        res->MoveNext();
    }

    res=0;


    wxString sql_pk = wxT("select pg_constraint.conname as pk_name,pg_attribute.attname as colname,pg_type.typname as typename from  \
                                    pg_constraint  inner join pg_class  \
                                    on pg_constraint.conrelid = pg_class.oid  \
                                    inner join pg_attribute on pg_attribute.attrelid = pg_class.oid  \
                                    and  pg_attribute.attnum = pg_constraint.conkey[1] \
                                    inner join pg_type on pg_type.oid = pg_attribute.atttypid \
                                    where pg_class.relname = '")+ table_name + wxT("' and pg_constraint.contype='p';");


    wxPostgreSQLresult  * res_tmp = wxGetApp().app_sql_select(sql_pk);

    if(res_tmp)
    {
        primay_keys.Clear();
        int irows = res_tmp->GetRowsNumber();
        res_tmp->MoveFirst();
        for(int i=0; i<irows; i++)
        {
            primay_keys.Add(res_tmp->GetVal(wxT("colname")));
            res_tmp->MoveNext();
        }
    }
    res_tmp->Clear();
/*
    if(isEditable)
    {

    }*/
}

sqlResultTable::~sqlResultTable()
{
    //dtor
    res=0;

    delete [] row_attr;

}

wxGridCellAttr *sqlResultTable::GetAttr(int row, int col, wxGridCellAttr::wxAttrKind  kind)
{
   if(m_data.IsEmpty())
   {
		wxGridCellAttr *attr = new wxGridCellAttr(row_attr[row].attr);
		attr->SetReadOnly();
		return attr;

   }else
   {
       row_attr[row].attr->IncRef();
       return row_attr[row].attr;
   }
}

void sqlResultTable::storedata(int irow, int icol)
{
    wxString str_clause,str_value;
    int icount = primay_keys.GetCount();

    switch(colTypes.Item(icol))
    {
    case PGOID_TYPE_BOOL:
        if(m_data[irow][icol] ==wxT("Y"))
            str_value = wxT("TRUE");
        else str_value = wxT("FALSE");
        break;
    case PGOID_TYPE_INT8:
    case PGOID_TYPE_SERIAL8:
    case PGOID_TYPE_INT2:
    case PGOID_TYPE_INT4:
        if(colNames.Item(icol)== wxT("project_catalog"))
        {
            if(m_data[irow][icol] ==wxT("Common Project"))
                str_value=wxT("1");

            if(m_data[irow][icol]==wxT("High-Speed Project"))
                str_value=wxT("2");

            if(m_data[irow][icol]==wxT("Special Project"))
                str_value=wxT("3");

            if(m_data[irow][icol]==wxT("Major Project"))
                str_value=wxT("4");

            if(m_data[irow][icol]==wxT("Pre-engineering"))
                str_value=wxT("5");

            if(m_data[irow][icol]==wxT("JY"))
                str_value=wxT("6");


        }
        else if(colNames.Item(icol)== wxT("nonstd_level"))
       {
            if(m_data[irow][icol] ==wxT("Complete Standard"))
                str_value=wxT("0");

            if(m_data[irow][icol] ==wxT("Option Standard"))
                str_value=wxT("1");

            if(m_data[irow][icol] ==wxT("Standard Configuration Without PC"))
                str_value=wxT("2");

            if(m_data[irow][icol] == wxT("Simple Non-Standard"))
                str_value=wxT("3");

            if(m_data[irow][icol] == wxT("Complex Non-Standard"))
                str_value=wxT("4");
       }
        else
                str_value = m_data[irow][icol];

        break;
    case PGOID_TYPE_DATE:
    case PGOID_TYPE_TIME:
    case PGOID_TYPE_TIMETZ:
    case PGOID_TYPE_TIMESTAMP:
    case PGOID_TYPE_TIMESTAMPTZ:
    case PGOID_TYPE_INTERVAL:
        str_value = m_data[irow][icol] + wxT(" 23:00:00");
        break;
    default:
        str_value = m_data[irow][icol];
        break;
    }

    if(icount==0)
        return;

    int i=0;
    wxArrayString stra_value = GetPrimaryKeyValue(irow);
    do
    {
        str_clause = str_clause+ primay_keys.Item(i)+wxT("='")+stra_value.Item(i)+wxT("' ");
        if(i!=icount-1)
            str_clause = str_clause+ wxT(" AND ");

        i++;
    }
    while (i<icount);
    wxString sql_query = wxT("Update ")+ table_name +wxT(" SET ")+colNames.Item(icol)+wxT("= '")+str_value+ wxT("' ")
                         wxT(" WHERE ")+str_clause+wxT(" ; ");
    if(!wxGetApp().app_sql_update(sql_query))
    {
        str_value = wxString::Format(_T("%d行,"),irow+1)+colNames.Item(icol)+_T("更新失败");
        wxLogMessage(str_value);
        return;
    }
}

wxArrayString sqlResultTable::GetPrimaryKeyValue(int row)
{
    wxArrayString stra;

    int icount = primay_keys.GetCount();

    for(int i=0;i<icount;i++)
    {
        for(int j=0;j<cols_num;j++)
        {
            if(primay_keys.Item(i)==colNames.Item(j))
            {
               stra.Add(m_data[row][j]);
               break;
            }
        }
    }

    return stra;

}

wxString sqlResultTable::GetValue(int row, int col)
{
    return m_data[row][col];
}

void sqlResultTable::SetValue(int row, int col, const wxString &value)
{
    m_data[row][col] = value;
    if(b_Save)
      storedata(row,col);
}


wxString sqlResultTable::GetColLabelValue (int col)
{
    return col_Labels.Item(col);
}

bool sqlResultTable::AppendRows (size_t numRows)
{
    /*
   wxGrid *grid = GetView();

   if (grid != NULL)
   {
      const int iNumRecords = GetNumberRows();
      const int iGridRows = grid->GetNumberRows();
      const int iNeedRows = iNumRecords - iGridRows;

      if (iNeedRows!=0)
      {
         grid->BeginBatch();
         grid->ClearSelection();

         if (grid->IsCellEditControlEnabled())
         {
            grid->DisableCellEditControl();
         }

         if(iNumRecords<iGridRows)
         {
            wxGridTableMessage pop(this,
                 wxGRIDTABLE_NOTIFY_ROWS_DELETED,
                 0, numRows);
            grid->ProcessTableMessage(pop);
         }

         if(iNumRecords>iGridRows)
         {
            wxGridTableMessage push(this,
                 wxGRIDTABLE_NOTIFY_ROWS_APPENDED,
                 iNumRecords);
            grid->ProcessTableMessage(push);
         }

         grid->AutoSize();
         grid->ForceRefresh();
         grid->EndBatch();
      }
   }

   return true;*/
    size_t curNumRows = m_data.GetCount();
    size_t curNumCols = ( curNumRows > 0
                         ? m_data[0].GetCount()
                         : ( GetView() ? GetView()->GetNumberCols() : 0 ) );

    wxArrayString sa;
    if ( curNumCols > 0 )
    {
        sa.Alloc( curNumCols );
        sa.Add( wxEmptyString, curNumCols );
    }

    m_data.Add( sa, numRows );

    if ( GetView() )
    {
        wxGridTableMessage msg( this,
                                wxGRIDTABLE_NOTIFY_ROWS_APPENDED,
                                numRows );

        GetView()->ProcessTableMessage( msg );
    }

    return true;

}

bool sqlResultTable::DeleteRows( size_t pos, size_t numRows)
{
        size_t curNumRows = m_data.GetCount();

    if ( pos >= curNumRows )
    {
        wxFAIL_MSG( wxString::Format
                    (
                        wxT("Called wxGridStringTable::DeleteRows(pos=%lu, N=%lu)\nPos value is invalid for present table with %lu rows"),
                        (unsigned long)pos,
                        (unsigned long)numRows,
                        (unsigned long)curNumRows
                    ) );

        return false;
    }

    if ( numRows > curNumRows - pos )
    {
        numRows = curNumRows - pos;
    }

    if ( numRows >= curNumRows )
    {
        m_data.Clear();
    }
    else
    {
        m_data.RemoveAt( pos, numRows );
    }

    if ( GetView() )
    {
        wxGridTableMessage msg( this,
                                wxGRIDTABLE_NOTIFY_ROWS_DELETED,
                                pos,
                                numRows );

        GetView()->ProcessTableMessage( msg );
    }

    return true;

}

bool sqlResultTable::InsertRows( size_t pos, size_t numRows)
{
    size_t curNumRows = m_data.GetCount();
    size_t curNumCols = ( curNumRows > 0 ? m_data[0].GetCount() :
                          ( GetView() ? GetView()->GetNumberCols() : 0 ) );

    if ( pos >= curNumRows )
    {
        return AppendRows( numRows );
    }

    wxArrayString sa;
    sa.Alloc( curNumCols );
    sa.Add( wxEmptyString, curNumCols );
    m_data.Insert( sa, pos, numRows );

    if ( GetView() )
    {
        wxGridTableMessage msg( this,
                                wxGRIDTABLE_NOTIFY_ROWS_INSERTED,
                                pos,
                                numRows );

        GetView()->ProcessTableMessage( msg );
    }

    return true;
}

void sqlResultTable::SetColLabelValue(int col, const wxString &label)
{
    return;
}

void sqlResultTable::SetRowLabelValue(int row, const wxString &label)
{
    return;
}

wxString sqlResultTable::GetRowLabelValue(int row)
{
    return row_Labels.Item(row);
}

void sqlResultTable::SetLabelValue(wxString label)
{
    col_Labels= wxStringTokenize(label, wxT(";"), wxTOKEN_DEFAULT );
}
