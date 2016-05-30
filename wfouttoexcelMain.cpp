/***************************************************************
 * Name:      wfouttoexcelMain.cpp
 * Purpose:   Code for Application Frame
 * Author:     ()
 * Created:   2014-03-27
 * Copyright:  ()
 * License:
 **************************************************************/
#include <locale.h>
#include "wfouttoexcelMain.h"
#include <wx/msgdlg.h>
#include <wx/dirdlg.h>
#include "pugixml.hpp"
#include "wfouttoexcelApp.h"
#include <wx/log.h>
#include "datepickerdlg.h"

//(*InternalHeaders(wfouttoexcelDialog)
#include <wx/intl.h>
#include <wx/string.h>
//*)

//helper functions
enum wxbuildinfoformat {
    short_f, long_f };

wxString wxbuildinfo(wxbuildinfoformat format)
{
    wxString wxbuild(wxVERSION_STRING);

    if (format == long_f )
    {
#if defined(__WXMSW__)
        wxbuild << _T("-Windows");
#elif defined(__UNIX__)
        wxbuild << _T("-Linux");
#endif

#if wxUSE_UNICODE
        wxbuild << _T("-Unicode build");
#else
        wxbuild << _T("-ANSI build");
#endif // wxUSE_UNICODE
    }

    return wxbuild;
}

//(*IdInit(wfouttoexcelDialog)
const long wfouttoexcelDialog::ID_TEXTCTRL1 = wxNewId();
const long wfouttoexcelDialog::ID_BUTTON1 = wxNewId();
const long wfouttoexcelDialog::ID_BUTTON2 = wxNewId();
const long wfouttoexcelDialog::ID_BUTTON4 = wxNewId();
const long wfouttoexcelDialog::ID_BUTTON5 = wxNewId();
const long wfouttoexcelDialog::ID_BUTTON8 = wxNewId();
const long wfouttoexcelDialog::ID_BUTTON_CONF = wxNewId();
const long wfouttoexcelDialog::ID_BUTTON6 = wxNewId();
const long wfouttoexcelDialog::ID_BUTTON_CONF_NSTD = wxNewId();
const long wfouttoexcelDialog::ID_BUTTON_NSTD_REASON = wxNewId();
const long wfouttoexcelDialog::ID_BUTTON7 = wxNewId();
const long wfouttoexcelDialog::ID_BUTTON3 = wxNewId();
const long wfouttoexcelDialog::ID_BUTTON_REVIEW = wxNewId();
const long wfouttoexcelDialog::ID_BUTTON_REVIEW_RESTART_HISTORY = wxNewId();
const long wfouttoexcelDialog::ID_BUTTON_BOM_CHANGE = wxNewId();
const long wfouttoexcelDialog::ID_GAUGE_STEP = wxNewId();
const long wfouttoexcelDialog::ID_TEXTCTRL2 = wxNewId();
//*)

BEGIN_EVENT_TABLE(wfouttoexcelDialog,wxDialog)
    //(*EventTable(wfouttoexcelDialog)
    //*)
END_EVENT_TABLE()

wfouttoexcelDialog::wfouttoexcelDialog(wxWindow* parent,wxWindowID id)
{
    //(*Initialize(wfouttoexcelDialog)
    wxBoxSizer* BoxSizer4;
    wxStaticBoxSizer* StaticBoxSizer2;
    wxBoxSizer* BoxSizer5;
    wxBoxSizer* BoxSizer2;
    wxBoxSizer* BoxSizer1;
    wxStaticBoxSizer* StaticBoxSizer1;
    wxBoxSizer* BoxSizer3;

    Create(parent, wxID_ANY, _("wxWidgets app"), wxDefaultPosition, wxDefaultSize, wxDEFAULT_DIALOG_STYLE, _T("wxID_ANY"));
    BoxSizer1 = new wxBoxSizer(wxVERTICAL);
    StaticBoxSizer1 = new wxStaticBoxSizer(wxHORIZONTAL, this, _("导出文件路径"));
    TextCtrl1 = new wxTextCtrl(this, ID_TEXTCTRL1, wxEmptyString, wxDefaultPosition, wxDefaultSize, wxTE_READONLY, wxDefaultValidator, _T("ID_TEXTCTRL1"));
    StaticBoxSizer1->Add(TextCtrl1, 6, wxALL|wxEXPAND, 0);
    Button1 = new wxButton(this, ID_BUTTON1, _("..."), wxDefaultPosition, wxDefaultSize, 0, wxDefaultValidator, _T("ID_BUTTON1"));
    StaticBoxSizer1->Add(Button1, 1, wxALL|wxEXPAND, 0);
    BoxSizer1->Add(StaticBoxSizer1, 1, wxALL|wxEXPAND, 0);
    BoxSizer2 = new wxBoxSizer(wxHORIZONTAL);
    Button2 = new wxButton(this, ID_BUTTON2, _("导出"), wxDefaultPosition, wxDefaultSize, 0, wxDefaultValidator, _T("ID_BUTTON2"));
    BoxSizer2->Add(Button2, 1, wxALL|wxEXPAND, 0);
    Button4 = new wxButton(this, ID_BUTTON4, _("导出所有项目"), wxDefaultPosition, wxDefaultSize, 0, wxDefaultValidator, _T("ID_BUTTON4"));
    BoxSizer2->Add(Button4, 1, wxALL|wxEXPAND, 0);
    Button5 = new wxButton(this, ID_BUTTON5, _("导出所有已完成项目"), wxDefaultPosition, wxDefaultSize, 0, wxDefaultValidator, _T("ID_BUTTON5"));
    BoxSizer2->Add(Button5, 1, wxALL|wxEXPAND, 0);
    Button7 = new wxButton(this, ID_BUTTON8, _("导出重排产"), wxDefaultPosition, wxDefaultSize, 0, wxDefaultValidator, _T("ID_BUTTON8"));
    BoxSizer2->Add(Button7, 1, wxALL|wxEXPAND, 0);
    Button_Conf = new wxButton(this, ID_BUTTON_CONF, _("导出配置任务表"), wxDefaultPosition, wxDefaultSize, 0, wxDefaultValidator, _T("ID_BUTTON_CONF"));
    BoxSizer2->Add(Button_Conf, 1, wxALL|wxEXPAND, 0);
    BoxSizer1->Add(BoxSizer2, 1, wxALL|wxEXPAND, 0);
    BoxSizer4 = new wxBoxSizer(wxHORIZONTAL);
    Button6 = new wxButton(this, ID_BUTTON6, _("导出非标追踪表"), wxDefaultPosition, wxDefaultSize, 0, wxDefaultValidator, _T("ID_BUTTON6"));
    BoxSizer4->Add(Button6, 1, wxALL|wxEXPAND, 0);
    Button_Conf_Nstd = new wxButton(this, ID_BUTTON_CONF_NSTD, _("导出非标申请处理表"), wxDefaultPosition, wxDefaultSize, 0, wxDefaultValidator, _T("ID_BUTTON_CONF_NSTD"));
    BoxSizer4->Add(Button_Conf_Nstd, 1, wxALL|wxEXPAND, 0);
    Button_Reason = new wxButton(this, ID_BUTTON_NSTD_REASON, _("非标分类导出"), wxDefaultPosition, wxDefaultSize, 0, wxDefaultValidator, _T("ID_BUTTON_NSTD_REASON"));
    BoxSizer4->Add(Button_Reason, 1, wxALL|wxEXPAND, 0);
    button_work = new wxButton(this, ID_BUTTON7, _("导出日常任务"), wxDefaultPosition, wxDefaultSize, 0, wxDefaultValidator, _T("ID_BUTTON7"));
    BoxSizer4->Add(button_work, 1, wxALL|wxEXPAND, 0);
    Button3 = new wxButton(this, ID_BUTTON3, _("退出"), wxDefaultPosition, wxDefaultSize, 0, wxDefaultValidator, _T("ID_BUTTON3"));
    BoxSizer4->Add(Button3, 1, wxALL|wxEXPAND, 0);
    BoxSizer1->Add(BoxSizer4, 1, wxALL|wxEXPAND, 0);
    BoxSizer5 = new wxBoxSizer(wxHORIZONTAL);
    Button_Review = new wxButton(this, ID_BUTTON_REVIEW, _("导出评审记录"), wxDefaultPosition, wxSize(129,50), 0, wxDefaultValidator, _T("ID_BUTTON_REVIEW"));
    BoxSizer5->Add(Button_Review, 1, wxALL|wxEXPAND, 0);
    Button_Review_his = new wxButton(this, ID_BUTTON_REVIEW_RESTART_HISTORY, _("导出评审重启历史记录"), wxDefaultPosition, wxDefaultSize, 0, wxDefaultValidator, _T("ID_BUTTON_REVIEW_RESTART_HISTORY"));
    BoxSizer5->Add(Button_Review_his, 1, wxALL|wxEXPAND, 0);
    Button_bom_change = new wxButton(this, ID_BUTTON_BOM_CHANGE, _("导出BOM更改通知单"), wxDefaultPosition, wxDefaultSize, 0, wxDefaultValidator, _T("ID_BUTTON_BOM_CHANGE"));
    BoxSizer5->Add(Button_bom_change, 1, wxALL|wxEXPAND, 0);
    BoxSizer1->Add(BoxSizer5, 1, wxALL, 0);
    StaticBoxSizer2 = new wxStaticBoxSizer(wxHORIZONTAL, this, _("导出进度"));
    gauge_process = new wxGauge(this, ID_GAUGE_STEP, 100, wxDefaultPosition, wxDefaultSize, 0, wxDefaultValidator, _T("ID_GAUGE_STEP"));
    StaticBoxSizer2->Add(gauge_process, 1, wxALL|wxEXPAND, 0);
    BoxSizer1->Add(StaticBoxSizer2, 1, wxALL|wxEXPAND, 0);
    BoxSizer3 = new wxBoxSizer(wxHORIZONTAL);
    tc_log = new wxTextCtrl(this, ID_TEXTCTRL2, wxEmptyString, wxDefaultPosition, wxDefaultSize, wxTE_MULTILINE, wxDefaultValidator, _T("ID_TEXTCTRL2"));
    BoxSizer3->Add(tc_log, 2, wxALL|wxEXPAND, 0);
    BoxSizer1->Add(BoxSizer3, 3, wxALL|wxEXPAND, 5);
    SetSizer(BoxSizer1);
    BoxSizer1->Fit(this);
    BoxSizer1->SetSizeHints(this);

    Connect(ID_BUTTON1,wxEVT_COMMAND_BUTTON_CLICKED,(wxObjectEventFunction)&wfouttoexcelDialog::OnButton1Click);
    Connect(ID_BUTTON2,wxEVT_COMMAND_BUTTON_CLICKED,(wxObjectEventFunction)&wfouttoexcelDialog::OnButton2Click);
    Connect(ID_BUTTON4,wxEVT_COMMAND_BUTTON_CLICKED,(wxObjectEventFunction)&wfouttoexcelDialog::OnButton4Click);
    Connect(ID_BUTTON5,wxEVT_COMMAND_BUTTON_CLICKED,(wxObjectEventFunction)&wfouttoexcelDialog::OnButton5Click);
    Connect(ID_BUTTON8,wxEVT_COMMAND_BUTTON_CLICKED,(wxObjectEventFunction)&wfouttoexcelDialog::OnButton7Click);
    Connect(ID_BUTTON_CONF,wxEVT_COMMAND_BUTTON_CLICKED,(wxObjectEventFunction)&wfouttoexcelDialog::OnButton_ConfClick);
    Connect(ID_BUTTON6,wxEVT_COMMAND_BUTTON_CLICKED,(wxObjectEventFunction)&wfouttoexcelDialog::OnButton6Click);
    Connect(ID_BUTTON_CONF_NSTD,wxEVT_COMMAND_BUTTON_CLICKED,(wxObjectEventFunction)&wfouttoexcelDialog::OnButton_Conf_NstdClick);
    Connect(ID_BUTTON_NSTD_REASON,wxEVT_COMMAND_BUTTON_CLICKED,(wxObjectEventFunction)&wfouttoexcelDialog::OnButton_ReasonClick);
    Connect(ID_BUTTON7,wxEVT_COMMAND_BUTTON_CLICKED,(wxObjectEventFunction)&wfouttoexcelDialog::Onbutton_workClick);
    Connect(ID_BUTTON3,wxEVT_COMMAND_BUTTON_CLICKED,(wxObjectEventFunction)&wfouttoexcelDialog::OnButton3Click);
    Connect(ID_BUTTON_REVIEW,wxEVT_COMMAND_BUTTON_CLICKED,(wxObjectEventFunction)&wfouttoexcelDialog::OnButton_ReviewClick);
    Connect(ID_BUTTON_REVIEW_RESTART_HISTORY,wxEVT_COMMAND_BUTTON_CLICKED,(wxObjectEventFunction)&wfouttoexcelDialog::OnButton_Review_hisClick);
    Connect(ID_BUTTON_BOM_CHANGE,wxEVT_COMMAND_BUTTON_CLICKED,(wxObjectEventFunction)&wfouttoexcelDialog::OnButton_bom_changeClick);
    //*)


    m_file = wxDateTime::Today().Format("%Y%m%d");

    pugi::xml_document ldoc;
    pugi::xml_node lnode;
    pugi::xml_parse_result result;

    result = ldoc.load_file("config.xml");
    wxString str_path;
    if(result)
    {
        lnode =ldoc.child("export");
        str_path = wxString(lnode.child_value("path"));
    }

    if(!str_path.IsEmpty())
        TextCtrl1->SetValue(str_path+wxT("\\")+m_file);


#if wxUSE_LOG
#ifdef __WXMOTIF__
    // For some reason, we get a memcpy crash in wxLogStream::DoLogStream
    // on gcc/wxMotif, if we use wxLogTextCtl. Maybe it's just gcc?
    delete wxLog::SetActiveTarget(new wxLogStderr);
#else
    // set our text control as the log target
    wxLogTextCtrl *logWindow = new wxLogTextCtrl(tc_log);
    delete wxLog::SetActiveTarget(logWindow);
#endif
#endif // wxUSE_LOG
}

wfouttoexcelDialog::~wfouttoexcelDialog()
{
    //(*Destroy(wfouttoexcelDialog)
    //*)

#if wxUSE_LOG
    delete wxLog::SetActiveTarget(NULL);
#endif // wxUSE_LOG
}

void wfouttoexcelDialog::OnQuit(wxCommandEvent& event)
{
    Close();
}

void wfouttoexcelDialog::OnAbout(wxCommandEvent& event)
{
    wxString msg = wxbuildinfo(long_f);
    wxMessageBox(msg, _("Welcome to..."));
}

void wfouttoexcelDialog::OnButton2Click(wxCommandEvent& event)
{
    init_array_head();
    wxString str_file;
    str_file = TextCtrl1->GetValue()+wxT(".xls");

    if(str_file.IsEmpty())
        return;

    wxLogMessage(_("导出文件:")+str_file);

    wxDateTime dt_now = wxDateTime::Today();
/*    wxString strSql = wxT("SELECT * ,(select doc_desc from s_doc where doc_id = step_desc_id) as remarks from v_task_list1 WHERE req_delivery_date >='")+
                      DateToAnsiStr(dt_now)+wxT("' AND workflow_id = 'WF0002' ORDER BY instance_id, flow_ser ASC");*/
    wxString strSql = wxT("SELECT *  from v_task_out1 WHERE (req_delivery_date >='")+
                      DateToAnsiStr(dt_now)+wxT("' or project_catalog = 6) AND workflow_id = 'WF0002' OR workflow_id='WF0006' ORDER BY instance_id, flow_ser ASC");

    export_excel(strSql, str_file);
}

void wfouttoexcelDialog::export_excel(wxString str_sql, wxString str_file)
{
    wxPostgreSQLresult *  res = wxGetApp().app_sql_select(str_sql);


    if(res->Status()!= PGRES_TUPLES_OK)
    {
        res->Clear();
        return;
    }

    int i_count = res->GetRowsNumber();


    int i_unit=0;

    bool b_continue;
    wxString temp_wbs;

    workbook wb1;
    worksheet* ws = wb1.sheet("Sheet1");


    int i_flow_ser;
    s_struct * result;
    int i ;
    for(i = 0;i<35;i++)
    {
        ws->label(0,i,array_head.Item(i).ToStdWstring());
    }

    res->MoveFirst();
    result = read_res(res);
    b_continue = false;

    //add by workflow_id='WF0006'
    int i_temp_ser=-1;
    wxString str1,str2,str3;
    int i_check=0;
    //add by workflow_id='WF0006'

    gauge_process->SetRange(i_count);
    for(i=0;i<i_count;i++)
    {
        i_flow_ser = result->i_flow_ser;
 //       i_total_flow = result->i_total_flow;

        gauge_process->SetValue(i+1);
        if(temp_wbs != result->array_value.Item(1))
        {
            i_unit++;
            b_continue = false;
            temp_wbs = result->array_value.Item(1);
            i_temp_ser=-1;

            label_result(i_unit,ws,result);

            wxLogMessage(_("输出")+temp_wbs);
        }

        if(i_temp_ser == i_flow_ser && temp_wbs==result->array_value.Item(1))
        {
              result->array_value.Item(11)=str1+"/"+result->array_value.Item(11);
              result->array_value.Item(12)=str2+"/"+result->array_value.Item(12);

              if(!result->array_value.Item(13).IsEmpty())
                   result->array_value.Item(13)=str3+"/"+result->array_value.Item(13);

              if(i_flow_ser<4 && b_continue && i_check<1)
              {
                  b_continue=false;
                  i_check++;
              }else
                  i_check=0;


        }

        i_temp_ser=i_flow_ser;

        if(b_continue)
        {
             res->MoveNext();
            delete result;
            result = read_res(res);
            continue;
        }

        if(temp_wbs == result->array_value.Item(1))
        {
            if(result->b_active )
                b_continue = true;

            ws->label(i_unit,11+(i_flow_ser-1)*3, result->array_value.Item(11).ToStdWstring());
            ws->label(i_unit,11+(i_flow_ser-1)*3+1, result->array_value.Item(12).ToStdString());
            ws->label(i_unit,11+(i_flow_ser-1)*3+2, result->array_value.Item(13).ToStdString());
        }
        str1= result->array_value.Item(11);
        str2=result->array_value.Item(12);
        str3=result->array_value.Item(13);

        res->MoveNext();
        delete result;
        result = read_res(res);
    }

    std::string filename = str_file.ToStdString();
    int error_msg = wb1.Dump(filename);

    if(error_msg == NO_ERRORS)
    {
            wxLogMessage(_("共计输出：")+NumToStr(i_unit));
    }


    res->Clear();
    gauge_process->SetValue(0);

}

void wfouttoexcelDialog::label_result(int irow,worksheet* _ws, s_struct* s_res)
{
    for(int i = 0; i<12;i++)
    {
        if(array_format.Item(i)== 0)
        {
             _ws->label(irow, i, s_res->array_value.Item(i).ToStdString());
        }

        if(array_format.Item(i)== 1)
        {
            _ws->label(irow,i, s_res->array_value.Item(i).ToStdWstring());
        }

        if(array_format.Item(i) == 5)
        {
            _ws->label(irow,i, s_res->array_value.Item(i).ToStdString());
        }

        if(array_format.Item(i)==8)
        {
            _ws->label(irow, i, s_res->array_value.Item(i).ToStdString());
        }
    }

    for(int i = 14;i< 16;i++)
    {
        _ws->label(irow,18+i, s_res->array_value.Item(i).ToStdString());

    }

    _ws->label(irow, 34, s_res->array_value.Item(16).ToStdString());
}

void wfouttoexcelDialog::OnButton3Click(wxCommandEvent& event)
{
    Close();
}

void wfouttoexcelDialog::OnButton1Click(wxCommandEvent& event)
{
    wxDirDialog dlg(NULL, "请选择文件输出路径", "",  wxDD_DEFAULT_STYLE | wxDD_DIR_MUST_EXIST);

    if (dlg.ShowModal() == wxID_CANCEL)
            return;

    wxString str_path;

    str_path = dlg.GetPath();

    TextCtrl1->SetValue(str_path+wxT("\\")+m_file);

    pugi::xml_document ldoc;
    pugi::xml_node lnode;
    pugi::xml_parse_result result;

    result = ldoc.load_file("config.xml");
    if(result)
    {
        lnode =ldoc.child("export").child("path");
        lnode.text().set(str_path.mbc_str());


        ldoc.save_file("config.xml");
    }
}

s_struct* wfouttoexcelDialog::read_res(wxPostgreSQLresult *_res)
{
    s_struct* wf_result = new s_struct;
    wf_result->b_active = _res->GetBool(wxT("is_active"));
    wf_result->i_flow_ser = _res->GetInt(wxT("flow_ser"));
    wf_result->i_total_flow = _res->GetInt(wxT("total_flow"));
    wf_result->s_action_id = _res->GetVal(wxT("action_id"));
    wf_result->s_workflow_id = _res->GetVal(wxT("workflow_id"));
    wxString str;

    int i_type;
    bool b_urgent;

    for(int i = 0;i< 17; i++)
    {
        if(i == 14)
        {
            int i_value  = _res->GetInt(array_str.Item(32));
            str = NumToStr(i_value);
        }else if(i == 15)
        {
            double  d_value = _res->GetDouble(array_str.Item(33));
            str = NumToStr(d_value);
        }
        else if(i == 6)
        {
            i_type = _res->GetInt(array_str.Item(i));
            switch(i_type)
            {
            case 1:
                str = wxT("Common Project");
                break;
            case 2:
                str = wxT("High-Speed Project");
                break;
            case 3:
                str = wxT("Special Project");
                break;
            case 4:
                str = wxT("Major Project");
                break;
            case 5:
                str = wxT("Pre_engineering");
                break;
            case 6:
                str = wxT("Lean Project");
            default:
                break;
            }
        }else if(i == 7)
        {
            b_urgent = _res->GetBool(array_str.Item(i));
            if(b_urgent)
                str = wxT("Y");
            else str.Empty();
        } else if(i ==8 )
        {
            i_type = _res->GetInt(array_str.Item(i));
            switch(i_type)
            {
            case -1:
                str=wxT("已删除");
                break;
            case 0:
                str=wxT("新项目");
                break;

            case 1:
                str=wxT("流程流转状态");
                break;
            case 2:
                str= wxT("评审完成");
                break;
            case 3:
                str= wxT("重排产");
                break;
            case 4:
                str=wxT("已暂停");
                break;
            case 5:
                str= wxT("配置完成");
                break;
            case 6:
                str= wxT("CM已经release");
                break;
            case 8:
                str = wxT("发运完成");
                break;
            default:
                break;
            }
        }else if(array_format.Item(i)== 4 )
        {
            wxDateTime dt_temp = _res->GetDate(array_str.Item(i));
            str = DateToAnsiStr(dt_temp);

        }else
        {
            str = _res->GetVal(array_str.Item(i));
        }

        if(i==16)
        {
            str = nstd_level_to_str(_res->GetInt(array_str.Item(34)));
        }

        wf_result->array_value.Add(str);
    }

    return wf_result;
}
//0-文本， 1-中文文本， 2-数字(int), 3 数字（double) ，4- bool, 5 -date , 6 -公式， 7-超链接

void wfouttoexcelDialog::init_array_head()
{
    wxString str_head, str;
    int i_format;
    for(int i=0;i<35;i++)
    {
       switch(i)
       {
     case 0:
          str_head = _("合同号");
          str = wxT("contract_id");
          i_format = 0;
          break;
     case 1:
          str_head = _("wbs_no");
          str=wxT("instance_id");
          i_format = 0;
          break;
     case 2:
          str_head = _("项目名称");
          str= wxT("project_name");
          i_format = 1;
          break;
     case 3:
          str_head = _("梯号");
          str = wxT("lift_no");
          i_format = 1;
          break;
     case 4:
          str_head = _("梯型编号");
          str =  wxT("elevator_id");
          i_format = 0;
          break;
     case 5:
          str_head = _("梯型");
          str = wxT("elevator_type");
          i_format = 0;
          break;
     case 6:
          str_head = _("项目类型");
          str = wxT("project_catalog");
          i_format = 0;
          break;
     case 7:
          str_head = _("是否紧急");
          str = wxT("is_urgent");
          i_format = 0;
          break;
     case 8:
          str_head = _("状态");
          str = wxT("status");
          i_format = 1;
          break;
     case 9:
          str_head = _("发货期");
          str = wxT("req_delivery_date");
          i_format = 5;
          break;
     case 10:
          str_head = _("要求配置完成日期");
          str = wxT("req_configure_finish");
          i_format = 5;
          break;
     case 11:
          str_head = _("项目授权—负责人");
          str = wxT("name");
          i_format = 1;
          break;
     case 12:
          str_head = _("项目授权—起始日期");
          str = wxT("start_date");
          i_format = 5;
          break;
     case 13:
          str_head = _("项目授权-完成日期");
          str = wxT("finish_date");
          i_format = 5;
          break;
     case 14:
          str_head = _("项目配置—负责人");
          str = wxT("name");
          i_format = 1;
          break;
     case 15:
          str_head = _("项目配置—起始日期");
          str = wxT("start_date");
          i_format = 5;
          break;
     case 16:
          str_head = _("项目配置-完成日期");
          str = wxT("finish_date");
          i_format = 5;
          break;
     case 17:
          str_head = _("配置校对—负责人");
          str = wxT("name");
          i_format = 1;
          break;
     case 18:
          str_head = _("配置校对—起始日期");
          str = wxT("start_date");
          i_format = 5;
          break;
     case 19:
          str_head = _("配置校对-完成日期");
          str = wxT("finish_date");
          i_format = 5;
          break;
     case 20:
          str_head = _("配置审核—负责人");
          str = wxT("name");
          i_format = 1;
          break;
     case 21:
          str_head = _("配置审核—起始日期");
          str = wxT("start_date");
          i_format = 5;
          break;
     case 22:
          str_head = _("配置审核-完成日期");
          str = wxT("finish_date");
          i_format = 5;
          break;
     case 23:
          str_head = _("接口传输—负责人");
          str = wxT("name");
          i_format = 1;
          break;
     case 24:
          str_head = _("接口传输—起始日期");
          str = wxT("start_date");
          i_format = 5;
          break;
     case 25:
          str_head = _("接口传输-完成日期");
          str = wxT("finish_date");
          i_format = 5;
          break;

    case 26:
          str_head = _("分箱任务分配—负责人");
          str = wxT("name");
          i_format = 1;
          break;
     case 27:
          str_head = _("分箱任务分配—起始日期");
          str = wxT("start_date");
          i_format = 5;
          break;
     case 28:
          str_head = _("分箱任务分配-完成日期");
          str = wxT("finish_date");
          i_format = 5;
          break;
     case 29:
          str_head = _("项目分箱—负责人");
          str = wxT("name");
          i_format = 1;
          break;
     case 30:
          str_head = _("项目分箱—起始日期");
          str = wxT("start_date");
          i_format = 5;
          break;
     case 31:
          str_head = _("项目分箱-完成日期");
          str = wxT("finish_date");
          i_format = 5;
          break;
     case 32:
          str_head = _("载重");
          str = wxT("load");
          i_format = 0;
          break;
     case 33:
          str_head = _("速度");
          str = wxT("speed");
          i_format = 0;
          break;
     case 34:
          str_head = _("非标等级");
          str = wxT("nonstd_level");
          i_format = 0;
          break;
     default:
          str_head.Empty();
          str.Empty();
          i_format = 0;
          break;

       }
       array_head.Add(str_head);
       array_str.Add(str);
       array_format.Add(i_format);
    }
}

void wfouttoexcelDialog::OnButton4Click(wxCommandEvent& event)
{
    init_array_head();

    wxString str_file;
    str_file = TextCtrl1->GetValue()+wxT(".xls");

    if(str_file.IsEmpty())
        return;

    wxLogMessage(_("导出文件:")+str_file);

    wxString strSql = wxT("SELECT *  from v_task_out1 WHERE workflow_id = 'WF0002' or workflow_id='WF0006' ORDER BY instance_id, flow_ser ASC");

    export_excel(strSql, str_file);
}

void wfouttoexcelDialog::OnButton5Click(wxCommandEvent& event)
{
    init_array_head();

    wxString str_file;
    str_file = TextCtrl1->GetValue()+wxT(".xls");

    if(str_file.IsEmpty())
        return;

    wxLogMessage(_("导出文件:")+str_file);
    datepickerdlg dlg;

    if(dlg.ShowModal()==wxID_CANCEL)
    {
        return;
    }

    bool b_from = dlg.b_from;
    bool b_to = dlg.b_to;

    wxDateTime dt_from,dt_to;

    if(b_from)
        dt_from = dlg.dt_from;

    if(b_to)
        dt_to = dlg.dt_to;


    wxString strSql,strSql1;

    wxArrayString a_blank_wbs;


    if(b_from&&b_to)
    {
        strSql1 = wxT("SELECT *  from v_task_out1 WHERE status >= 5 AND (workflow_id = 'WF0002' OR workflow_id='WF0006') and req_configure_finish>='")+DateToStrFormat(dt_from)+wxT(" 00:00:00' and req_configure_finish<='")+DateToStrFormat(dt_to)+wxT(" 23:59:00'  ORDER BY instance_id, flow_ser ASC");
        strSql = wxT("SELECT instance_id  from v_task_out1 WHERE status >= 5 AND action_id = 'AT00000008' and req_configure_finish is NULL AND finish_date >='")+DateToStrFormat(dt_from)+wxT(" 00:00:00' and finish_date<='")+DateToStrFormat(dt_to)+wxT(" 23:59:00'  ORDER BY instance_id, flow_ser ASC");
    }

    if(!b_from&&b_to)
    {

        strSql1 = wxT("SELECT *  from v_task_out1 WHERE status >= 5 AND (workflow_id = 'WF0002' OR workflow_id='WF0006') and req_configure_finish<='")+DateToStrFormat(dt_to)+wxT(" 23:59:00'  ORDER BY instance_id, flow_ser ASC");
        strSql = wxT("SELECT instance_id  from v_task_out1 WHERE status >= 5 AND action_id = 'AT00000008' and req_configure_finish is NULL and finish_date<='")+DateToStrFormat(dt_to)+wxT(" 23:59:00'  ORDER BY instance_id, flow_ser ASC");
    }

    if(b_from&&!b_to)
    {
        strSql1 = wxT("SELECT *  from v_task_out1 WHERE status >= 5 AND (workflow_id = 'WF0002' OR workflow_id='WF0006') and req_configure_finish>='")+DateToStrFormat(dt_from)+wxT(" 00:00:00' ORDER BY instance_id, flow_ser ASC");
         strSql = wxT("SELECT instance_id  from v_task_out1 WHERE status >= 5 AND action_id = 'AT00000008' and req_configure_finish is NULL and finish_date>='")+DateToStrFormat(dt_from)+wxT(" 00:00:00' ORDER BY instance_id, flow_ser ASC");
    }


    if(!b_from&&!b_to)
    {
        strSql = wxT("SELECT instance_id  from v_task_out1 WHERE status >= 5 AND (workflow_id = 'WF0002' OR workflow_id='WF0006') ORDER BY instance_id, flow_ser ASC");
       // strSql2.Empty();
    }else
    {
        wxPostgreSQLresult *  res = wxGetApp().app_sql_select(strSql);

        if(res->Status()!= PGRES_TUPLES_OK)
        {
            res->Clear();
            wxLogMessage(_("精益项目查找失败，SQL语句错误!"));
            strSql = strSql1;
            goto export_sec;
        }

        int i_count = res->GetRowsNumber();

        if(i_count==0)
        {
            strSql = strSql1;
            res->Clear();
            goto export_sec;
        }

        int i=0;
        for(i=0;i<i_count;i++)
        {
            wxString str = res->GetVal(wxT("instance_id"));

            a_blank_wbs.Add(str);

            res->MoveNext();
        }


        res->Clear();

        //strSql.Empty();
        strSql= wxT("SELECT *  from v_task_out1 WHERE (workflow_id = 'WF0002' OR workflow_id='WF0006') and (");/* ORDER BY instance_id, flow_ser ASC");*/


        for(i=0;i<i_count;i++)
        {
            if(i==i_count-1)
                strSql = strSql + wxT("instance_id='")+a_blank_wbs.Item(i)+wxT("' )");
            else
                strSql = strSql + wxT("instance_id='")+a_blank_wbs.Item(i)+wxT("' or ");

        }

        strSql = strSql+ wxT(" order by instance_id, flow_ser asc ");

        strSql = wxT("(")+strSql1+wxT(") union all (")+strSql+wxT(") ");

        a_blank_wbs.Clear();

    }

    export_sec:
       export_excel(strSql, str_file);
}

void wfouttoexcelDialog::OnButton6Click(wxCommandEvent& event)
{
    wxString str_file;
    str_file = TextCtrl1->GetValue()+wxT("-FB.xls");

    if(str_file.IsEmpty())
        return;

    wxLogMessage(_("导出文件:")+str_file);

    init_nonstd_array_head();

    export_nonstd_excel(str_file);
}

//0-文本， 1-中文文本， 2-数字(int), 3 数字（double) ，4- bool, 5 -date , 6 -公式， 7-超链接
void wfouttoexcelDialog::init_nonstd_array_head()
{
    array_head.Clear();
    wxString str_head;
    for(int i=0;i<22;i++)
    {
       switch(i)
       {
     case 0:
          str_head = _("编号");
          break;
     case 1:
          str_head = _("非标申请书");
          break;
     case 2:
          str_head = _("是否有图纸");
          break;
     case 3:
          str_head = _("合同号");
          break;
     case 4:
          str_head = _("项目名称");
          break;
     case 5:
          str_head = _("WBS号");
          break;
    case 6:
          str_head = _("梯号");
          break;
     case 7:
          str_head = _("申请人");
          break;
     case 8:
          str_head = _("物料申请表接收时间");
          break;
     case 9:
          str_head = _("PDM pending");
          break;
     case 10:
          str_head = _("PDM ACTIVITE");
          break;
     case 11:
          str_head = _("IE流转文件");
          break;
     case 12:
          str_head = _("SAP READY");
          break;
     case 13:
          str_head = _("资料室接收时间");
          break;
     case 14:
          str_head = _("资料室完成时间");
          break;
     case 15:
           str_head = _("计划物料交付日期");
           break;
     case 16:
           str_head = _("计划图纸交付日期");
           break;
     case 17:
           str_head = _("是否有非标物料");
           break;
     case 18:
           str_head = _("非标分类");
           break;
     case 19:
           str_head = _("非标原因");
           break;
     case 20:
           str_head = _("分任务非标原因");
           break;
     case 21:
           str_head = _("分任务备注");
           break;
     default:
          str_head.Empty();
          break;

       }
       array_head.Add(str_head);
    }
}

void wfouttoexcelDialog::export_nonstd_excel(wxString str_file)
{
    datepickerdlg dlg;

    if(dlg.ShowModal()==wxID_CANCEL)
    {
        return;
    }

    bool b_from = dlg.b_from;
    bool b_to = dlg.b_to;

    wxDateTime dt_from,dt_to;

    if(b_from)
        dt_from = dlg.dt_from;

    if(b_to)
        dt_to = dlg.dt_to;

    wxString str_sql = wxT("select nstd_mat_app_id,index_mat_id,has_nonstd_draw,contract_id,project_id,project_name,flow_mask, link_list,\
                           res_person,res_engineer, (select name from s_employee where employee_id = res_person) as res_person_name,\
                           (select name from s_employee where employee_id = res_engineer) as res_engineer_name, mat_req_date, drawing_req_date, has_nonstd_mat,\
                           nonstd_catalog,nonstd_desc,instance_nstd_desc,instance_remarks from v_nonstd_app_item_instance where nstd_mat_app_id != '' AND status !='-1' ");

    if(b_from  && !b_to)
    {
        str_sql = str_sql + wxT("and create_date >='")+DateToStrFormat(dt_from)+wxT("' ORDER BY nstd_mat_app_id ASC;");

    }else if (!b_from && b_to)
    {
        str_sql = str_sql + wxT("and create_date <='")+DateToStrFormat(dt_to)+wxT("' ORDER BY nstd_mat_app_id ASC;");

    }else if(b_from && b_to)
    {
        str_sql = str_sql + wxT("and create_date <='")+DateToStrFormat(dt_to)+wxT("' and create_date>='")+DateToStrFormat(dt_from)+wxT("' ORDER BY nstd_mat_app_id ASC;");
    }else
    {
        str_sql = str_sql + wxT(" ORDER BY nstd_mat_app_id ASC;");

    }

    wxPostgreSQLresult * res = wxGetApp().app_sql_select(str_sql);

    if(res->Status()!= PGRES_TUPLES_OK)
    {
        res->Clear();
        return;
    }

    workbook wb1;
    worksheet* ws = wb1.sheet("Sheet1");

    int i_flow_ser;
    s_struct * result;
    int i ;

    for(i = 0;i<22;i++)
    {
        ws->label(0,i,array_head.Item(i).ToStdWstring());
    }

    int i_count = res->GetRowsNumber();
    res->MoveFirst();
    wxString str_nstd_id, str, str_res_employee;
//操作代码
    gauge_process->SetRange(i_count);
    for(int i=0; i<i_count; i++)
    {
        gauge_process->SetValue(i+1);
        str_nstd_id  = res->GetVal(wxT("nstd_mat_app_id"));

        ws->label(i+1, 0, str_nstd_id.ToStdString());

        str = res->GetVal(wxT("index_mat_id"));

        ws->label(i+1, 1, str.ToStdString());

        bool b_draw = res->GetBool(wxT("has_nonstd_draw"));

        if(b_draw)
            str = wxT("Y");
        else str.Empty();

        ws->label(i+1, 2, str.ToStdString());

        str = res->GetVal(wxT("contract_id"));

        ws->label(i+1, 3, str.ToStdString());

        str = res->GetVal(wxT("project_name"));

        ws->label(i+1, 4, str.ToStdWstring());

        str = res->GetVal(wxT("link_list"));
        ws->label(i+1, 5, str.ToStdString());


        wxString s_project_id = res->GetVal(wxT("project_id"));
        wxString s_flow_mask = res->GetVal(wxT("flow_mask"));

        if(s_project_id == wxT("CS") || s_project_id == wxT("OTHER")|| s_project_id == wxT("MOD") || s_flow_mask == wxT("000"))
        {
            str_res_employee = res->GetVal(wxT("res_person_name"));
        }else
        {
          str_res_employee = res->GetVal(wxT("res_engineer_name"));
        }

        if(!str_res_employee.IsEmpty() && str_res_employee != wxT("''"))
            ws->label(i+1, 7, str_res_employee.ToStdWstring());

        str=DateToStrFormat(res->GetDate(wxT("mat_req_date")));
        ws->label(i+1, 15, str.ToStdWstring());

        str=DateToStrFormat(res->GetDate(wxT("drawing_req_date")));
        ws->label(i+1, 16, str.ToStdWstring());

        str = res->GetVal(wxT("nonstd_catalog"));
        ws->label(i+1, 18, str.ToStdWstring());

        str = res->GetVal(wxT("nonstd_desc"));
        ws->label(i+1, 19, str.ToStdWstring());

        str = res->GetVal(wxT("instance_nstd_desc"));
        ws->label(i+1, 20, str.ToStdWstring());

        str= res->GetVal(wxT("instance_remarks"));
        ws->label(i+1,21, str.ToStdWstring());

        b_draw = res->GetBool(wxT("has_nonstd_mat"));

        if(b_draw)
            str = wxT("Y");
        else str.Empty();

        ws->label(i+1, 17, str.ToStdString());

        wxString str_index_mat_id  = res->GetVal(wxT("index_mat_id"));

        str_sql = wxT("select * from v_task_list3 where instance_id = '")+str_index_mat_id+wxT("' AND (workflow_id = 'WF0004' OR workflow_id = 'WF0005') ORDER BY workflow_id, flow_ser ASC;");

        wxPostgreSQLresult * low_res = wxGetApp().app_sql_select(str_sql);

        if(low_res->Status()!= PGRES_TUPLES_OK)
        {
            low_res->Clear();
            return;
        }

        int i_low_count = low_res->GetRowsNumber();

        wxString s_action_id, s_workflow_id;

        for(int j=0;j<i_low_count;j++)
        {
              s_action_id = low_res->GetVal(wxT("action_id"));
              s_workflow_id = low_res->GetVal(wxT("workflow_id"));

              if(s_workflow_id == wxT("WF0004") && s_action_id == wxT("AT00000015"))
              {
                  str = low_res->GetVal(wxT("start_date"));
                  ws->label(i+1, 13, str.ToStdString());

                  str = low_res->GetVal(wxT("finish_date"));
                  ws->label(i+1,14, str.ToStdString());
              }
/*
              if(s_workflow_id == wxT("WF0005") && s_action_id == wxT("AT00000016"))
              {
                  str = low_res->GetVal(wxT("name"));
                  ws->label(i+1, 7, str.ToStdWstring());

              }*/

              if(s_workflow_id == wxT("WF0005") && s_action_id == wxT("AT00000018"))
              {
                str = low_res->GetVal(wxT("start_date"));
                  ws->label(i+1, 8, str.ToStdString());

                  str = low_res->GetVal(wxT("finish_date"));
                  ws->label(i+1, 9, str.ToStdString());
              }

              if(s_workflow_id == wxT("WF0005") && s_action_id == wxT("AT00000019"))
              {
                  str = low_res->GetVal(wxT("finish_date"));
                  ws->label(i+1, 10, str.ToStdString());
              }

              if(s_workflow_id == wxT("WF0005") && s_action_id == wxT("AT00000020"))
              {
                  str = low_res->GetVal(wxT("finish_date"));
                  ws->label(i+1, 12, str.ToStdString());
              }

              low_res->MoveNext();
        }

        low_res->Clear();

        res->MoveNext();

    }

//操作代码

    std::string filename = str_file.ToStdString();
    int error_msg = wb1.Dump(filename);

    if(error_msg == NO_ERRORS)
    {
            wxLogMessage(_("共计输出：")+NumToStr(i_count));
    }


    res->Clear();
    gauge_process->SetValue(0);
}

void wfouttoexcelDialog::Onbutton_workClick(wxCommandEvent& event)
{
    wxString str_file;
    str_file = TextCtrl1->GetValue()+wxT("-daily.xls");

    if(str_file.IsEmpty())
        return;

    wxLogMessage(_("导出文件:")+str_file);

    init_daily_work_head();

    export_daily_work(str_file);
}

void wfouttoexcelDialog::init_daily_work_head()
{
    array_head.Clear();
    wxString str_head;
    for(int i=0;i<10;i++)
    {
       switch(i)
       {
     case 0:
          str_head = _("编号");
          break;
     case 1:
          str_head = _("任务分类");
          break;
     case 2:
          str_head = _("任务内容");
          break;
     case 3:
          str_head = _("任务量(小时)");
          break;
     case 4:
          str_head = _("工作日(from)");
          break;
     case 5:
          str_head = _("工作日(to)");
          break;

     case 6:
          str_head = _("Approved by");
          break;
     case 7:
          str_head = _("Approved Date");
          break;

    case 8:
          str_head = _("提交人");
          break;
     case 9:
          str_head = _("提交日期");
          break;
     default:
          str_head.Empty();
          break;

       }
       array_head.Add(str_head);
    }

}

void wfouttoexcelDialog::export_daily_work(wxString str_file)
{
    wxString str_sql = wxT("SELECT task_id,task_catalog,task_content,review_load,submiter,submit_time,reviewer,review_time,\
                           work_date,work_date_to, (select name from s_employee where employee_id = submiter) as submit_name,\
                           (select name from s_employee where employee_id = reviewer) as review_name from l_daily_task where task_status = 2; ");

    wxPostgreSQLresult * res = wxGetApp().app_sql_select(str_sql);

    if(res->Status()!= PGRES_TUPLES_OK)
    {
        res->Clear();
        return;
    }


    workbook wb1;
    worksheet* ws = wb1.sheet("Sheet1");

    int i_flow_ser;
    s_struct * result;
    int i ;

    for(i = 0;i<10;i++)
    {
        ws->label(0,i,array_head.Item(i).ToStdWstring());
    }

    int i_count = res->GetRowsNumber();

    gauge_process->SetRange(i_count);
    res->MoveFirst();
    wxString str;
    wxDateTime dt_temp;
//操作代码
    for(int i=0; i<i_count; i++)
    {
        gauge_process->SetValue(i+1);
        str  = res->GetVal(wxT("task_id"));

        ws->label(i+1, 0, str.ToStdString());

        str = res->GetVal(wxT("task_catalog"));

        ws->label(i+1, 1, str.ToStdWstring());

        str = res->GetVal(wxT("task_content"));

        ws->label(i+1, 2, str.ToStdWstring());

        double d_load= res->GetDouble(wxT("review_load"));

        ws->number(i+1, 3, d_load);

        str = res->GetVal(wxT("work_date"));

       // str = DateToStrFormat(dt_temp);

        ws->label(i+1, 4, str.ToStdWstring());

        str = res->GetVal(wxT("work_date_to"));
       // str = DateToStrFormat(dt_temp);
        ws->label(i+1, 5, str.ToStdString());

        str = res->GetVal(wxT("submit_name"));
        ws->label(i+1, 8, str.ToStdWstring());

        dt_temp = res->GetDate(wxT("submit_time"));
        str = DateToStrFormat(dt_temp);
        ws->label(i+1, 9, str.ToStdString());

        str = res->GetVal(wxT("review_name"));
        ws->label(i+1, 6, str.ToStdWstring());


        dt_temp = res->GetDate(wxT("review_time"));
        str = DateToStrFormat(dt_temp);
        ws->label(i+1, 7, str.ToStdString());


        res->MoveNext();

    }

//操作代码

    std::string filename = str_file.ToStdString();
    int error_msg = wb1.Dump(filename);

    if(error_msg == NO_ERRORS)
    {
            wxLogMessage(_("共计输出：")+NumToStr(i_count));
    }


    res->Clear();
    gauge_process->SetValue(0);
}

void wfouttoexcelDialog::OnButton7Click(wxCommandEvent& event)
{
    wxString str_file;
    str_file = TextCtrl1->GetValue()+wxT("-restart.xls");

    if(str_file.IsEmpty())
        return;

    wxLogMessage(_("导出文件:")+str_file);


    export_restart_porject(str_file);
}

void wfouttoexcelDialog::init_restart_project_head()
{
    array_head.Clear();
    wxString str_head;
    for(int i=0;i<15;i++)
    {
       switch(i)
       {
     case 0:
          str_head = _("合同号");
          break;
     case 1:
          str_head = _("wbs no.");
          break;
     case 2:
          str_head = _("项目名称");
          break;
     case 3:
          str_head = _("梯号");
          break;
     case 4:
          str_head = _("梯型");
          break;
     case 5:
          str_head = _("载重");
          break;

     case 6:
          str_head = _("速度");
          break;

     case 7:
           str_head =_("非标等级");
          break;
     case 8:
          str_head = _("项目状态");
          break;
     case 9:
          str_head = _("最新一次排产时间");
          break;

    case 10:
          str_head = _("最新一次分箱完成时间");
          break;
     case 11:
          str_head = _("第一次排产时间");
          break;
     case 12:
          str_head = _("第一次分箱完成时间");
          break;
     case 13:
          str_head = _("第二次排产时间");
          break;
     case 14:
          str_head = _("第二次分箱完成时间");
          break;
     default:
          str_head.Empty();
          break;

       }
       array_head.Add(str_head);
    }


}

void wfouttoexcelDialog::export_restart_porject(wxString str_file)
{
    datepickerdlg dlg;

    if(dlg.ShowModal()==wxID_CANCEL)
    {
        return;
    }

    bool b_from = dlg.b_from;
    bool b_to = dlg.b_to;

    wxDateTime dt_from,dt_to;

    if(b_from)
        dt_from = dlg.dt_from;

    if(b_to)
        dt_to = dlg.dt_to;


    wxString str_sql, str;
    wxArrayString array_units;

    if(b_from && b_to)
        str_sql = wxT("select * from v_task_out1 where  is_restart= true and action_id = 'AT00000003' and start_date <='")+DateToAnsiStr(dt_to)+wxT("' and start_date >='")+DateToAnsiStr(dt_from)+wxT("';");

    if(b_from && !b_to)
        str_sql = wxT("select * from v_task_out1 where  is_restart= true and action_id = 'AT00000003' and start_date >='")+DateToAnsiStr(dt_from)+wxT("';");

    if(!b_from && b_to)
        str_sql = wxT("select * from v_task_out1 where  is_restart= true and action_id = 'AT00000003' and start_date <='")+DateToAnsiStr(dt_to)+wxT("';");

    wxPostgreSQLresult * res = wxGetApp().app_sql_select(str_sql);

    if(res->Status()!= PGRES_TUPLES_OK)
    {
        res->Clear();
        return;
    }

    int i_count = res->GetRowsNumber();
    res->MoveFirst();

    workbook wb1;
    worksheet* ws = wb1.sheet("Sheet1");

    gauge_process->SetRange(i_count);
    init_restart_project_head();

    for(int i = 0;i<15;i++)
    {
        ws->label(0,i,array_head.Item(i).ToStdWstring());
    }

    for(int i=0;i<i_count;i++)
    {
        gauge_process->SetValue(i+1);
        str = res->GetVal(wxT("instance_id"));

        wxLogMessage(str);

        label_newest_restart_info(i+1,str,ws);
        label_restart_info(i+1,str,ws);

        res->MoveNext();

    }

    res->Clear();

    std::string filename = str_file.ToStdString();
    int error_msg = wb1.Dump(filename);

    if(error_msg == NO_ERRORS)
    {
            wxLogMessage(_("共计输出：")+NumToStr(i_count));
    }

    gauge_process->SetValue(0);


}

void wfouttoexcelDialog::label_restart_info(int i_row, wxString str, worksheet* ws)
{
    wxString str_sql = wxT("select * from l_restart_log where instance_id ='")+str+wxT("' and (action_id = 'AT00000003' or action_id ='AT00000008') order by instance_id, restart_times, flow_ser asc;");

    wxPostgreSQLresult * res = wxGetApp().app_sql_select(str_sql);

    if(res->Status()!= PGRES_TUPLES_OK)
    {
        res->Clear();
        return;
    }

    int i_count = res->GetRowsNumber();

    wxString str_action_id;

    res->MoveFirst();

    wxDateTime dt_temp;
    int i_restart_times;
    for(int i=0;i<i_count;i++)
    {
        str_action_id = res->GetVal(wxT("action_id"));
        i_restart_times = res->GetInt(wxT("restart_times"));


        if(str_action_id == wxT("AT00000003"))
        {
            dt_temp = res->GetDateTime(wxT("start_date"));
            str = DateToStrFormat(dt_temp);
            ws->label(i_row, 11+i, str.ToStdString());

        }

        if(str_action_id == wxT("AT00000008"))
        {

            dt_temp = res->GetDateTime(wxT("finish_date"));
            str = DateToStrFormat(dt_temp);
            ws->label(i_row, 11+i, str.ToStdString());
        }

        res->MoveNext();
    }

    res->Clear();
}

void wfouttoexcelDialog::label_newest_restart_info(int i_row, wxString str, worksheet* ws)
{
    wxString str_sql = wxT("select * from v_task_out1 where instance_id ='")+str+wxT("' and (action_id = 'AT00000003' or action_id ='AT00000008') order by flow_ser asc;");

    wxPostgreSQLresult * res = wxGetApp().app_sql_select(str_sql);

    if(res->Status()!= PGRES_TUPLES_OK)
    {
        res->Clear();
        return;
    }

    int i_count = res->GetRowsNumber();

    res->MoveFirst();

    wxString str_instance;
    str = res->GetVal(wxT("contract_id"));

    ws->label(i_row, 0, str.ToStdString());

    str_instance = res->GetVal(wxT("instance_id"));
    ws->label(i_row, 1, str_instance.ToStdString());

    str = res->GetVal(wxT("project_name"));
    ws->label(i_row, 2, str.ToStdWstring());

    str = res->GetVal(wxT("lift_no"));
 //   str.Replace(wxT("*"), wxT(""));
    ws->label(i_row, 3, str.ToStdWstring());

    str = res->GetVal(wxT("elevator_type"));
    ws->label(i_row, 4, str.ToStdString());

    int i_value = res->GetInt(wxT("load"));
    ws->number(i_row, 5, i_value);

    double f_speed = res->GetDouble(wxT("speed"));
    ws->number(i_row, 6, f_speed);

    int i_nstd= res->GetInt(wxT("nonstd_level"));
    str = nstd_level_to_str(i_nstd);
    ws->label(i_row,7, str.ToStdString());


    int i_status = res->GetInt(wxT("status"));
    ws->label(i_row,8, prj_status_to_str(i_status).ToStdString());

    wxDateTime dt_temp = res->GetDateTime(wxT("start_date"));
    str = DateToStrFormat(dt_temp);
    ws->label(i_row, 9, str.ToStdString());


    if(i_count <= 1)
    {
        res->Clear();
        return;
    }


    res->MoveNext();

    if(str_instance == res->GetVal(wxT("instance_id"))&&i_status>=5)
    {

        dt_temp = res->GetDateTime(wxT("finish_date"));
        str = DateToStrFormat(dt_temp);
        ws->label(i_row, 10, str.ToStdString());
    }

    res->Clear();

}

void wfouttoexcelDialog::OnButton_ReasonClick(wxCommandEvent& event)
{
    wxString str_file;
    str_file = TextCtrl1->GetValue()+wxT("-worktime.xls");

    if(str_file.IsEmpty())
        return;

    wxLogMessage(_("导出文件:")+str_file);

    init_nonstd_catalog_head();

    export_nonstd_catalog_excel(str_file);
}

void wfouttoexcelDialog::init_nonstd_catalog_head()
{
    array_head.Clear();
    wxString str_head;
    for(int i=0;i<18;i++)
    {
       switch(i)
       {
     case 0:
          str_head = _("编号");
          break;
     case 1:
          str_head = _("非标申请书");
          break;
     case 2:
          str_head = _("是否有图纸");
          break;
     case 3:
           str_head = _("是否有非标物料");
           break;
     case 4:
           str_head = _("非标类别");
           break;
     case 5:
           str_head = _("非标原因");
           break;
     case 6:
           str_head = _("申请人");
           break;
     case 7:
           str_head = _("非标一级任务负责人");
           break;
     case 8:
           str_head = _("子任务负责人");
           break;
     case 9:
           str_head = _("确认工时");
           break;
     case 10:
           str_head = _("子任务收到时间");
           break;
     case 11:
           str_head = _("子任务完成时间");
           break;
     case 12:
           str_head = _("非标物料提交时间");
           break;
     case 13:
           str_head = _("非标设计提交时间");
           break;
     case 14:
          str_head = _("合同号");
          break;
     case 15:
          str_head = _("项目名称");
          break;
     case 16:
          str_head = _("WBS号");
          break;
/*     case 18:
           str_head = _("非标物料工作流状态")；
           break;
     case 19:
           str_head = _("非标设计工作流状态");
           break;*/
     default:
          str_head.Empty();
          break;

       }
       array_head.Add(str_head);
    }
}

void wfouttoexcelDialog::export_nonstd_catalog_excel(wxString str_file)
{
    datepickerdlg dlg;

    if(dlg.ShowModal()==wxID_CANCEL)
    {
        return;
    }

    bool b_from = dlg.b_from;
    bool b_to = dlg.b_to;

    wxDateTime dt_from,dt_to;

    if(b_from)
        dt_from = dlg.dt_from;

    if(b_to)
        dt_to = dlg.dt_to;

    wxString str_sql = wxT("select nstd_mat_app_id,item_ser, index_mat_id,has_nonstd_draw,contract_id,project_id,project_name,flow_mask, link_list,\
                           res_person,res_engineer, (select name from s_employee where employee_id = res_person) as res_person_name,\
                           (select name from s_employee where employee_id = res_engineer) as res_engineer_name, mat_req_date, drawing_req_date, has_nonstd_mat, \
                           nonstd_catalog,ins_start_date, ins_finish_date,nonstd_desc,index_id from v_nonstd_app_item_instance where  status !='-1' ");

    if(b_from  && !b_to)
    {
        str_sql = str_sql + wxT("and create_date >='")+DateToStrFormat(dt_from)+wxT("' ORDER BY index_mat_id ASC;");

    }else if (!b_from && b_to)
    {
        str_sql = str_sql + wxT("and create_date <='")+DateToStrFormat(dt_to)+wxT("' ORDER BY index_mat_id ASC;");

    }else if(b_from && b_to)
    {
        str_sql = str_sql + wxT("and create_date <='")+DateToStrFormat(dt_to)+wxT("' and create_date>='")+DateToStrFormat(dt_from)+wxT("' ORDER BY index_mat_id ASC;");
    }else
    {
        str_sql = str_sql + wxT(" ORDER BY index_mat_id ASC;");

    }
    wxPostgreSQLresult * res = wxGetApp().app_sql_select(str_sql);

    if(res->Status()!= PGRES_TUPLES_OK)
    {
        res->Clear();
        return;
    }

    workbook wb1;
    worksheet* ws = wb1.sheet("Sheet1");

    int i_flow_ser;
    s_struct * result;
    int i ;

    for(i = 0;i<18;i++)
    {
        ws->label(0,i,array_head.Item(i).ToStdWstring());
    }

    int i_count = res->GetRowsNumber();
    res->MoveFirst();
    wxString str_nstd_id, str, str_res_employee, str_index_id, s_res_person;
    gauge_process->SetRange(i_count);
//操作代码
    wxString str_index_id_just;
    for(i=0; i<i_count; i++)
    {
        gauge_process->SetValue(i+1);
        str_nstd_id  = res->GetVal(wxT("nstd_mat_app_id"));

        str_index_id = res->GetVal(wxT("index_id"));

        ws->label(i+1, 0, str_nstd_id.ToStdString());

        str = res->GetVal(wxT("index_mat_id"));

        ws->label(i+1, 1, str.ToStdString());

        bool b_draw = res->GetBool(wxT("has_nonstd_draw"));

        if(b_draw)
            str = wxT("Y");
        else str.Empty();

        ws->label(i+1, 2, str.ToStdString());

        bool b_mat = res->GetBool(wxT("has_nonstd_mat"));

        if(b_mat)
            str = wxT("Y");
        else str.Empty();

        ws->label(i+1, 3, str.ToStdString());

        str = res->GetVal(wxT("nonstd_catalog"));

        ws->label(i+1, 4, str.ToStdWstring());

        str = res->GetVal(wxT("nonstd_desc"));

        ws->label(i+1, 5, str.ToStdWstring());

        str = res->GetVal(wxT("contract_id"));

        ws->label(i+1, 14, str.ToStdString());

        str = res->GetVal(wxT("project_name"));

        ws->label(i+1, 15, str.ToStdWstring());

        str = res->GetVal(wxT("link_list"));
        ws->label(i+1, 16, str.ToStdString());

        str = res->GetVal(wxT("res_person_name"));
        ws->label(i+1, 7, str.ToStdWstring());

        s_res_person = str;

        str = res->GetVal(wxT("res_engineer_name"));
        ws->label(i+1, 8, str.ToStdWstring());

        str=DateToStrFormat(res->GetDate(wxT("ins_start_date")));
        ws->label(i+1, 10, str.ToStdString());

        str=DateToStrFormat(res->GetDate(wxT("ins_finish_date")));
        ws->label(i+1, 11, str.ToStdString());

        wxString str_index_mat_id  = res->GetVal(wxT("index_mat_id"));

        double d_draw_eval, d_mat_eval, d_total;
         wxPostgreSQLresult * low_res;
         int i_low_count;
        wxString s_action_id;
        if(b_draw||b_mat)
        {
            str_sql = wxT("select start_date,action_id from v_task_list3 where instance_id = '")+str_index_mat_id+wxT("' AND (action_id = 'AT00000012' OR action_id = 'AT00000016' ) ORDER BY workflow_id, flow_ser ASC;");

            low_res = wxGetApp().app_sql_select(str_sql);

            if(low_res->Status()!= PGRES_TUPLES_OK)
            {
                low_res->Clear();
                return;
            }

            i_low_count = low_res->GetRowsNumber();

            for(int j=0; j<i_low_count; j++)
            {
                s_action_id = low_res->GetVal(wxT("action_id"));

                if(b_mat && s_action_id == wxT("AT00000012"))
                {
                    str = low_res->GetVal(wxT("start_date"));
                    ws->label(i+1, 12, str.ToStdString());

                }


                if(b_draw && s_action_id == wxT("AT00000016"))
                {
                    str = low_res->GetVal(wxT("start_date"));
                    ws->label(i+1, 13, str.ToStdString());

                }

                low_res->MoveNext();
            }

            low_res->Clear();

            str_sql = wxT("select task_id,eval_value,action_id from l_evaluation_instance where task_id = '")+str_index_mat_id+wxT("' and (action_id = 'AT00000012' OR action_id ='AT00000016') \
                          AND is_valid=true;");
            low_res = wxGetApp().app_sql_select(str_sql);

            if(low_res->Status()!= PGRES_TUPLES_OK)
            {
                low_res->Clear();
                return;
            }

            i_low_count = low_res->GetRowsNumber();


            if(i_low_count ==0)
                ws->number(i+1, 9, 0);

            for(int k=0; k<i_low_count; k++)
            {
                d_draw_eval = 0.0;
                d_mat_eval = 0.0;
                d_total = 0.0;
                s_action_id = low_res->GetVal(wxT("action_id"));

                if( s_action_id == wxT("AT00000012"))
                {
                    d_mat_eval = low_res->GetDouble(wxT("eval_value"));
                }


                if( s_action_id == wxT("AT00000016"))
                {
                    d_draw_eval = low_res->GetDouble(wxT("eval_value"));
                }

                d_total =d_mat_eval+d_draw_eval;
                ws->number(i+1, 9, d_total);

                low_res->MoveNext();
            }

            low_res->Clear();

        }
        else if(str_index_id != str_index_id_just)
        {
            str_index_id_just = str_index_id;
            str_sql = wxT("select task_id,eval_value,action_id from l_evaluation_instance where task_id = '")+str_index_id+wxT("' and action_id = 'AT00000022' \
                          AND is_valid=true;");
            low_res = wxGetApp().app_sql_select(str_sql);

            if(low_res->Status()!= PGRES_TUPLES_OK)
            {
                low_res->Clear();
                return;
            }

            i_low_count = low_res->GetRowsNumber();

            if(i_low_count ==0)
                ws->number(i+1, 9, 0);

            for(int k=0; k<i_low_count; k++)
            {
                d_total =0.0;
                s_action_id = low_res->GetVal(wxT("action_id"));

                if( s_action_id == wxT("AT00000022"))
                {
                    d_total = low_res->GetDouble(wxT("eval_value"));
                }

                ws->number(i+1, 9, d_total);

                low_res->MoveNext();
            }

            low_res->Clear();
        }else
            ws->number(i+1, 9, 0);

        str_sql = wxT(" select name from v_task_list2 where instance_id ='")+str_index_id+wxT("' and action_id = 'AT00000009';");
        low_res = wxGetApp().app_sql_select(str_sql);

        if(low_res->Status()!= PGRES_TUPLES_OK)
        {
            low_res->Clear();
            return;
        }

        i_low_count = low_res->GetRowsNumber();

        if(i_low_count!=0)
        {
            str = low_res->GetVal(wxT("name"));
            ws->label(i+1,6, str.ToStdWstring());
        }
        low_res->Clear();

        if(i_low_count ==0)
        {
             ws->label(i+1,6, s_res_person.ToStdWstring());
        }

        res->MoveNext();

    }



//操作代码

    std::string filename = str_file.ToStdString();
    int error_msg = wb1.Dump(filename);

    if(error_msg == NO_ERRORS)
    {
            wxLogMessage(_("共计输出：")+NumToStr(i_count));
    }


    res->Clear();
    gauge_process->SetValue(0);
}

wxString wfouttoexcelDialog::nstd_status_to_str(int i_status)
{
    wxString str;
    switch(i_status)
    {
       case -1 :
         str =  wxT("DEL");
         break;
       case 0:
         str =  wxT("CRTD");
         break;
       case 1:
          str =  wxT("SUBMIT");
          break;
       case 2:
         str =  wxT("APPROVE");
         break;
       case 3:
         str =  wxT("PROC");
         break;
       case 4:
         str =  wxT("FEEDBACK");
         break;
       case 5:
         str = wxT("PDM-READY");
         break;
       case 6:
         str = wxT("SAP-READY");
         break;
       case 7:
          str = wxT("DOC-READY");
          break;
       case 8:
          str = wxT("ALL-READY");
          break;
       case 10:
         str =  wxT("CLOSED");
         break;
       default:
           str.Empty();
         break;
    }
    return str;
}

wxString wfouttoexcelDialog::prj_status_to_str(int i_prj)
{
    wxString str;
    switch(i_prj)
    {
       case -2:
          str = wxT("DEL");
          break;
       case -1 :
         str =  wxT("CANCEL");
         break;
       case 0:
         str =  wxT("CRTD");
         break;
       case 1:
         str =  wxT("ACTIVE");
         break;
       case 2:
          str =  wxT("FININISH");
          break;
       case 3:
         str =  wxT("RESTART");
         break;
       case 4:
         str =  wxT("FREEZED");
         break;
       case 5:
         str = wxT("CLOSED");
         break;
       case 6:
         str = wxT("RELEASED");
         break;
       case 7:
         str = wxT("PART FINISH");
         break;
       case 8:
         str = wxT("DELIVERIED");
         break;
       default:
           str.Empty();
         break;
    }
    return str;
}

int wfouttoexcelDialog::prj_str_to_status(wxString str_status)
{
    if(str_status == wxT("DEL"))
        return -2;

    if(str_status == wxT("CANCEL"))
       return -1;

    if(str_status == wxT("CRTD"))
        return 0;

    if(str_status == wxT("ACTIVE"))
        return 1;

    if(str_status == wxT("FININISH"))
        return 2;

    if(str_status == wxT("RESTART"))
        return 3;

    if(str_status == wxT("FREEZED"))
        return 4;

    if(str_status == wxT("CLOSED"))
        return 5;

    if(str_status == wxT("RELEASED"))
        return 6;

    if(str_status == wxT("PART FINISH"));
        return 7;

    if(str_status == wxT("DELIVERIED"));
       return 8;

    return -2;
}

void wfouttoexcelDialog::OnButton_ConfClick(wxCommandEvent& event)
{
    wxString str_file;
    str_file = TextCtrl1->GetValue()+wxT("-conf_eval.xls");

    if(str_file.IsEmpty())
        return;

    wxLogMessage(_("导出文件:")+str_file);


    export_conf_evaluation(str_file);
}

void wfouttoexcelDialog::init_conf_evaluation()
{
   array_head.Clear();
    wxString str_head;
    for(int i=0;i<14;i++)
    {
       switch(i)
       {
     case 0:
          str_head = _("合同号");
          break;
     case 1:
          str_head = _("项目名称");
          break;
     case 2:
          str_head = _("wbs NO(连写)");
          break;
     case 3:
          str_head = _("台数");
          break;
     case 4:
          str_head = _("任务收到日期");
          break;
     case 5:
          str_head = _("任务完成日期");
          break;

     case 6:
          str_head = _("配置工程师");
          break;
     case 7:
          str_head = _("核查人");
          break;
     case 8:
          str_head = _("评估人");
          break;

    case 9:
          str_head = _("单梯配置平均耗时");
          break;
     case 10:
          str_head = _("配置耗时");
          break;
     default:
          str_head.Empty();
          break;

       }
       array_head.Add(str_head);
    }


}

void wfouttoexcelDialog::export_conf_evaluation(wxString str_file)
{
    datepickerdlg dlg;

    if(dlg.ShowModal()==wxID_CANCEL)
    {
        return;
    }

    bool b_from = dlg.b_from;
    bool b_to = dlg.b_to;

    wxDateTime dt_from,dt_to;

    if(b_from)
        dt_from = dlg.dt_from;

    if(b_to)
        dt_to = dlg.dt_to;



    wxString str_sql = wxT("select item_id,eval_id, task_id, task_desc, operator_id,(select name from s_employee where employee_id = operator_id) as operator_name, op_start_date, op_finish_date, \
                            evaluator_id,(select name from s_employee where employee_id = evaluator_id) as evaluator_name, eval_value, eval_remarks,error_qty, error_point, check_id, \
                            (select name from s_employee where employee_id = check_id) as check_name from l_evaluation_instance \
                            where is_valid = true and action_id = 'AT00000004' AND ");


    if(b_from && b_to)
        str_sql = str_sql + wxT(" op_finish_date <='")+DateToAnsiStr(dt_to)+wxT("' and op_finish_date >='")+DateToAnsiStr(dt_from)+wxT("';");

    if(b_from && !b_to)
        str_sql = str_sql + wxT(" op_finish_date >='")+DateToAnsiStr(dt_from)+wxT("';");

    if(!b_from && b_to)
        str_sql = str_sql + wxT(" op_finish_date <='")+DateToAnsiStr(dt_to)+wxT("';");

    wxPostgreSQLresult * res = wxGetApp().app_sql_select(str_sql);

    if(res->Status()!= PGRES_TUPLES_OK)
    {
        res->Clear();
        return;
    }

    workbook wb1;
    worksheet* ws = wb1.sheet("Sheet1");

    init_conf_evaluation();
    for(int i = 0;i<11;i++)
    {
        ws->label(0,i,array_head.Item(i).ToStdWstring());
    }

    int i_count = res->GetRowsNumber();
    wxArrayString ar_str;

    wxString str_task_id,str_project_id, str;
    int i_units;

    gauge_process->SetRange(i_count);

    for(int i=0;i<i_count;i++)
    {
         i_units = 0;
         str_task_id= res->GetVal(wxT("task_id"));
         str_project_id = str_task_id.Left(10);
         str = get_contract_id(str_project_id);

         ws->label(i+1,0,str.ToStdString());

         str = res->GetVal(wxT("task_desc"));

         ar_str = wxStringTokenize(str,wxT(":"),wxTOKEN_DEFAULT );

         if(!ar_str.IsEmpty())
         {
            ws->label(i+1,1,ar_str.Item(0).ToStdWstring());

            if(ar_str.GetCount()>=2)
            {

                str=ar_str.Item(1);
            }
         }else
            str.Empty();

        ar_str.Clear();

        if(!str.IsEmpty())
        {
            ar_str =  wxStringTokenize(str,wxT(","),wxTOKEN_RET_EMPTY );

            ar_str.Sort(CompareWBSString);

            str = combine_str(ar_str,wxT("~"),wxT(","));

            ws->label(i+1,2,str.ToStdString());

            i_units =ar_str.GetCount();


        }

        str=wxString::Format("%d",i_units);

        ws->label(i+1,3,str.ToStdString());

        str = res->GetVal(wxT("op_start_date"));

        ws->label(i+1,4,str.ToStdString());

        str = res->GetVal(wxT("op_finish_date"));

        ws->label(i+1,5,str.ToStdString());


        str = res->GetVal(wxT("operator_name"));
        ws->label(i+1,6,str.ToStdWstring());

        str = res->GetVal(wxT("check_name"));
        ws->label(i+1,7,str.ToStdWstring());

        str=res->GetVal(wxT("evaluator_name"));
        ws->label(i+1,8,str.ToStdWstring());

        double f_value = res->GetDouble(wxT("eval_value"));

        str = NumToStr(f_value);
        ws->label(i+1,9,str.ToStdString());

        double f_total = f_value * i_units;

        str = NumToStr(f_total);
        ws->label(i+1,10,str.ToStdString());

        str = res->GetVal(wxT("eval_id"));
        ws->label(i+1,11,str.ToStdString());

        gauge_process->SetValue(i);
        res->MoveNext();
    }

//操作代码

    std::string filename = str_file.ToStdString();
    int error_msg = wb1.Dump(filename);

    if(error_msg == NO_ERRORS)
    {
            wxLogMessage(_("共计输出：")+NumToStr(i_count));
    }


    res->Clear();

    gauge_process->SetValue(0);
}


wxString wfouttoexcelDialog::get_contract_id(wxString s_project_id)
{
    wxString str_sql = wxT("select contract_id from s_project_info where project_id = '")+s_project_id+wxT("';");
    wxPostgreSQLresult * res = wxGetApp().app_sql_select(str_sql);
    if(res->Status()!= PGRES_TUPLES_OK)
    {
        res->Clear();
        return wxEmptyString;
    }

    int i_count = res->GetRowsNumber();

    if(i_count ==0)
    {
        res->Clear();
        return wxEmptyString;
    }

    res->MoveFirst();

    wxString str = res->GetVal(wxT("contract_id"));

    res->Clear();

    return str;
}

wxString combine_str(wxArrayString &a_str,wxString serial_str, wxString individual_str)
{
    int i_count = a_str.GetCount();
    int i;
    wxString str,str1,str2,str_head;
    str_head = a_str.Item(0).Left(10);
    str= str1 = a_str.Item(0).Right(3);
    bool b_serial=false;
    for (i=1; i<i_count; i++)
    {
        str2=a_str.Item(i).Right(3);

        if(wxAtoi(str2)==wxAtoi(str1)+1)
        {
            str1=str2;
            b_serial=true;
        }else
        {
            if(b_serial)
            {
                str=str+serial_str+str1+individual_str+str2;
                b_serial = false;
            }
            else
                str =str+individual_str+str2;

            str1=str2;
        }
    }
    if(b_serial)
        str= str_head+wxT(".")+str+serial_str+ str1;
    else
       str = str_head+wxT(".")+str;

    return str;
}

int CompareWBSString(const wxString& first, const wxString& second)
{
    int i_first = wxAtoi(first.Right(3));
    int i_second = wxAtoi(second.Right(3));

    return i_first-i_second;
}

void wfouttoexcelDialog::OnButton_Conf_NstdClick(wxCommandEvent& event)
{
    wxString str_file;
    str_file = TextCtrl1->GetValue()+wxT("-Conf_Nstd_app.xls");

    if(str_file.IsEmpty())
        return;

    wxLogMessage(_("导出文件:")+str_file);

    init_conf_app_deal();

    export_conf_app_deal(str_file);
}


void wfouttoexcelDialog::init_conf_app_deal()
{
    array_head.Clear();
    array_str.Clear();
    array_format.Clear();
    wxString str_head, str;
    int i_format;
    for(int i=0;i<33;i++)
    {
       switch(i)
       {
     case 0:
          str_head = _("非标申请号");
          str= wxT("index_id");
          i_format = 0;
          break;
     case 1:
          str_head = _("项目号");
          str= wxT("project_id");
          i_format = 0;
          break;
     case 2:
          str_head = _("合同号");
          str = wxT("contract_id");
          i_format = 0;
          break;
     case 3:
          str_head = _("项目名称");
          str = wxT("project_name");
          i_format = 1;
          break;
     case 4:
          str_head = _("关联wbs_no");
          str = wxT("link_list");
          i_format  =0;
          break;
     case 5:
          str_head = _("类别");
          str = wxT("item_catalog");
          i_format = 0;
          break;
    case 6:
          str_head = _("非标类别");
          str = wxT("nonstd_catalog");
          i_format = 1;
          break;
     case 7:
          str_head = _("非标原因");
          str = wxT("nonstd_desc");
          i_format = 1;
          break;
     case 8:
          str_head = _("备注");
          str = wxT("remarks");
          i_format = 1;
          break;
     case 9:
          str_head = _("申请人");
          str= wxT("appname");
          i_format = 1;
          break;
     case 10:
          str_head = _("物料需求日");
          str = wxT("mat_req_date");
          i_format = 8;
          break;
     case 11:
          str_head = _("图纸需求日");
          str = wxT("drawing_req_date");
          i_format = 8;
          break;
     case 12:
          str_head = _("非标负责人");
          str = wxT("resname");
          i_format = 1;
          break;

     case 13:
          str_head = _("状态");
          str = wxT("status");
          i_format = 0;
          break;
     case 14:
          str_head = _("流程状态");
          str = wxT("wf_status");
          i_format = 1;
          break;
//此步开始为步骤信息
    case 15:
          str_head = _("非标申请—负责人");
          str = wxT("name");
          i_format = 1;
          break;
     case 16:
          str_head = _("非标申请—起始日期");
          str = wxT("start_date");
          i_format = 5;
          break;
     case 17:
          str_head = _("非标申请-完成日期");
          str = wxT("finish_date");
          i_format = 5;
          break;
     case 18:
          str_head = _("非标申请审核—负责人");
          str = wxT("name");
          i_format = 1;
          break;
     case 19:
          str_head = _("非标申请审核—起始日期");
          str = wxT("start_date");
          i_format = 5;
          break;
     case 20:
          str_head = _("非标申请审核-完成日期");
          str = wxT("finish_date");
          i_format = 5;
          break;
     case 21:
          str_head = _("非标任务分配—负责人");
          str = wxT("name");
          i_format = 1;
          break;
     case 22:
          str_head = _("非标任务分配—起始日期");
          str = wxT("start_date");
          i_format = 5;
          break;
     case 23:
          str_head = _("非标任务分配-完成日期");
          str = wxT("finish_date");
          i_format = 5;
          break;
     case 24:
          str_head = _("非标任务处理—负责人");
          str = wxT("name");
          i_format = 1;
          break;
     case 25:
          str_head = _("非标任务处理—起始日期");
          str = wxT("start_date");
          i_format = 5;
          break;
     case 26:
          str_head = _("非标任务处理-完成日期");
          str = wxT("finish_date");
          i_format = 5;
          break;
    case 27:
          str_head = _("非标处理审核—负责人");
          str = wxT("name");
          i_format = 1;
          break;
     case 28:
          str_head = _("非标处理审核—起始日期");
          str = wxT("start_date");
          i_format = 5;
          break;
     case 29:
          str_head = _("非标处理审核-完成日期");
          str = wxT("finish_date");
          i_format = 5;
          break;
     case 30:
          str_head = _("确认收到回复—负责人");
          str = wxT("name");
          i_format = 1;
          break;
     case 31:
          str_head = _("确认收到回复—起始日期");
          str = wxT("start_date");
          i_format = 5;
          break;
     case 32:
          str_head = _("确认收到回复-完成日期");
          str = wxT("finish_date");
          i_format = 5;
          break;
     default:
          str_head.Empty();
          str.Empty();
          i_format = 0;
          break;

       }
       array_head.Add(str_head);
       array_str.Add(str);
       array_format.Add(i_format);
    }
}

void wfouttoexcelDialog::export_conf_app_deal(wxString str_file)
{
    datepickerdlg dlg;

    if(dlg.ShowModal()==wxID_CANCEL)
    {
        return;
    }

    bool b_from = dlg.b_from;
    bool b_to = dlg.b_to;

    wxDateTime dt_from,dt_to;

    if(b_from)
        dt_from = dlg.dt_from;

    if(b_to)
        dt_to = dlg.dt_to;

    wxString s_sql_att;

    if(b_from && b_to)
        s_sql_att =  wxT(" mat_req_date <='")+DateToAnsiStr(dt_to)+wxT("' and mat_req_date >='")+DateToAnsiStr(dt_from)+wxT("' ");

    if(b_from && !b_to)
        s_sql_att = wxT(" mat_req_date >='")+DateToAnsiStr(dt_from)+wxT("' ");

    if(!b_from && b_to)
        s_sql_att = wxT(" mat_req_date <='")+DateToAnsiStr(dt_to)+wxT("' ");


    wxString str_sql = wxT(" select *,(select name from s_employee where employee_id=v_task_list2.create_emp_id) as appname, \
                           (select name from s_employee where employee_id=res_person) as resname  from v_task_list2 where \
                            ")+s_sql_att+wxT(" order by instance_id, flow_ser asc; ");
    wxPostgreSQLresult *  res = wxGetApp().app_sql_select(str_sql);


    if(res->Status()!= PGRES_TUPLES_OK)
    {
        res->Clear();
        return;
    }

    int i_count = res->GetRowsNumber();


    int i_unit=0;

    bool b_continue;
    wxString temp_task_id;

    workbook wb1;
    worksheet* ws = wb1.sheet("Sheet1");


    int i_flow_ser;
    s_struct * result;
    int i ;
    for(i = 0;i<33;i++)
    {
        ws->label(0,i,array_head.Item(i).ToStdWstring());
    }

    res->MoveFirst();
    result = read_conf_nstd_app_res(res);
    b_continue = false;

    gauge_process->SetRange(i_count);
    for(i=0;i<i_count;i++)
    {
        i_flow_ser = result->i_flow_ser;
 //       i_total_flow = result->i_total_flow;

        gauge_process->SetValue(i+1);
        if(temp_task_id != result->array_value.Item(0))
        {
            i_unit++;
            b_continue = false;
            temp_task_id = result->array_value.Item(0);

            label_conf_nstd_app_result(i_unit,ws,result);

            wxLogMessage(_("输出")+temp_task_id);
        }


        if(b_continue)
        {
             res->MoveNext();
            delete result;
            result = read_conf_nstd_app_res(res);
            continue;
        }

        if(temp_task_id == result->array_value.Item(0))
        {
            if(result->b_active)
                b_continue = true;

            ws->label(i_unit,15+(i_flow_ser-1)*3, result->array_value.Item(15).ToStdWstring());
            ws->label(i_unit,15+(i_flow_ser-1)*3+1, result->array_value.Item(16).ToStdString());
            ws->label(i_unit,15+(i_flow_ser-1)*3+2, result->array_value.Item(17).ToStdString());
        }

        res->MoveNext();
        delete result;
        result = read_conf_nstd_app_res(res);
    }

    std::string filename = str_file.ToStdString();
    int error_msg = wb1.Dump(filename);

    if(error_msg == NO_ERRORS)
    {
            wxLogMessage(_("共计输出：")+NumToStr(i_unit));
    }


    res->Clear();
    gauge_process->SetValue(0);
}

void wfouttoexcelDialog::label_conf_nstd_app_result(int irow,worksheet* _ws, s_struct* s_res)
{
    for(int i = 0; i<15;i++)
    {
        if(array_format.Item(i)== 0)
        {
             _ws->label(irow, i, s_res->array_value.Item(i).ToStdString());
        }

        if(array_format.Item(i)== 1)
        {
            _ws->label(irow,i, s_res->array_value.Item(i).ToStdWstring());
        }

        if(array_format.Item(i) == 5)
        {
            _ws->label(irow,i, s_res->array_value.Item(i).ToStdString());
        }

        if(array_format.Item(i)== 8)
        {
            _ws->label(irow, i, s_res->array_value.Item(i).ToStdString());
        }
    }
}

s_struct* wfouttoexcelDialog::read_conf_nstd_app_res(wxPostgreSQLresult *_res)
{
    s_struct* wf_result = new s_struct;
    wf_result->b_active = _res->GetBool(wxT("is_active"));
    wf_result->i_flow_ser = _res->GetInt(wxT("flow_ser"));
    wf_result->i_total_flow = _res->GetInt(wxT("total_flow"));
    wf_result->s_action_id = _res->GetVal(wxT("action_id"));
    wf_result->s_workflow_id = _res->GetVal(wxT("workflow_id"));
    wxString str;
    int i_status;

    for(int i = 0;i< 18; i++)
    {

        if(i==13)
        {
            str = nstd_status_to_str(_res->GetInt(array_str.Item(i)));
        }else {
        switch(array_format.Item(i))
        {
        case 0:
        case 1:
            str = _res->GetVal(array_str.Item(i));
            break;
        case 5:
            str = DateToAnsiStr(_res->GetDateTime(array_str.Item(i)));
            break;
        case 8:
            str = DateToStrFormat(_res->GetDateTime(array_str.Item(i)));
            break;
        default:
            str.Empty();
            break;
        }
        }

        wf_result->array_value.Add(str);
    }

    return wf_result;
}

wxString nstd_level_to_str(int i_status)
{
    wxString str;
    switch(i_status)
    {

    case 1:
        str = wxT("Target STD");
        break;
    case 2:
        str = wxT("Option STD");
        break;
    case 3:
        str = wxT("Simple Non-STD");
        break;
    case 4:
        str = wxT("Complex Non-STD");
        break;
    case 5:
        str = wxT("Comp-Standard");
        break;
    case 6:
        str = wxT("Comp-Measurement");
        break;
    case 7:
        str = wxT("Comp-Configurable");
        break;
    case 11:
        str = wxT("Design Fault-Qty.");
        break;
    case 12:
        str = wxT("Design Fault-Spec.");
        break;
    case 13:
        str = wxT("Sales Order Fault-Qt");
        break;
    case 14:
        str = wxT("Sales Order Fault-Sp");
        break;
    case 15:
        str = wxT("Matl Pick Fault-Qty.");
        break;
    case 16:
        str = wxT("Matl Pick Fault-Spec");
        break;
    case 17:
        str = wxT("Packing Fault-Qty.");
        break;
    case 18:
        str = wxT("Packing Fault-Spec.");
        break;
    case 19:
        str = wxT("Logistic Fault.");
        break;
    case 21:
        str = wxT("Abnormal in logistic");
        break;
    case 22:
        str = wxT("ATI or ECR.");
        break;
    case 23:
        str = wxT("Others");
        break;
    default:
        str.Empty();
        break;
    }

    return str;
}


wxString urgent_level_to_str(int i_level)
{
    wxString str;
    switch(i_level)
    {

    case 1:
        str = wxT("普通");
        break;
    case 2:
        str = wxT("紧急");
        break;
    case 3:
        str = wxT("特紧急");
        break;
    default:
        str= wxT("普通");
        break;
    }

    return str;
}

void wfouttoexcelDialog::OnButton_ReviewClick(wxCommandEvent& event)
{
    wxString str_file;
    str_file = TextCtrl1->GetValue()+wxT("-review.xls");

    if(str_file.IsEmpty())
        return;

    wxLogMessage(_("导出文件:")+str_file);

    init_review_header();

    export_review(str_file);
}

void wfouttoexcelDialog::init_review_header()
{
    array_head.Clear();
    array_str.Clear();
    array_format.Clear();
    wxString str_head, str;
    int i_format;
    for(int i=0;i<33;i++)
    {
       switch(i)
       {
       case 0:
          str_head = _("任务号");
          str = wxT("review_task_id");
          i_format = 0;
          break;
     case 1:
          str_head = _("合同号");
          str= wxT("contract_id");
          i_format = 0;
          break;
     case 2:
          str_head = _("项目名称");
          str= wxT("project_name");
          i_format = 1;
          break;
     case 3:
          str_head = _("wbs no");
          str = wxT("wbs_no");
          i_format = 0;
          break;
     case 4:
          str_head = _("梯号");
          str = wxT("lift_no");
          i_format = 0;
          break;
     case 5:
          str_head = _("商务负责人");
          str = wxT("res_cm_name");
          i_format  =1;
          break;
     case 6:
          str_head = _("评审负责人");
          str = wxT("review_engineer_name");
          i_format = 1;
          break;
    case 7:
          str_head = _("紧急程度");
          str = wxT("urgent_level");
          i_format = 1;
          break;
     case 8:
          str_head = _("评审要求完成日期");
          str = wxT("require_review_date");
          i_format = 8;
          break;
     case 9:
          str_head = _("图纸套数");
          str = wxT("review_drawing_qty");
          i_format = 1;
          break;
     case 10:
          str_head = _("重审次数");
          str= wxT("restart_times");
          i_format = 1;
          break;
     case 11:
          str_head = _("价格表收到日期");
          str = wxT("price_list_receive");
          i_format = 8;
          break;
     case 12:
        str_head = _("状态");
        str = wxT("unit_status");
        i_format = 0;
        break;
     case 13:
        str_head = _("流程状态");
        str = wxT("unit_wf_status");
        i_format = 1;
        break;

     case 14:
        str_head = _("分公司");
        str= wxT("branch_name");
        i_format=1;
        break;

//此步开始为步骤信息
    case 15:
          str_head = _("合同立项—负责人");
          str = wxT("name");
          i_format = 1;
          break;
     case 16:
          str_head = _("合同立项—起始日期");
          str = wxT("start_date");
          i_format = 5;
          break;
     case 17:
          str_head = _("合同立项-完成日期");
          str = wxT("finish_date");
          i_format = 5;
          break;
     case 18:
          str_head = _("项目授权—负责人");
          str = wxT("name");
          i_format = 1;
          break;
     case 19:
          str_head = _("项目授权—起始日期");
          str = wxT("start_date");
          i_format = 5;
          break;
     case 20:
          str_head = _("项目授权-完成日期");
          str = wxT("finish_date");
          i_format = 5;
          break;
     case 21:
          str_head = _("项目评审—负责人");
          str = wxT("name");
          i_format = 1;
          break;
     case 22:
          str_head = _("项目评审—起始日期");
          str = wxT("start_date");
          i_format = 5;
          break;
     case 23:
          str_head = _("项目评审-完成日期");
          str = wxT("finish_date");
          i_format = 5;
          break;
     case 24:
          str_head = _(" 评审审核—负责人");
          str = wxT("name");
          i_format = 1;
          break;
     case 25:
          str_head = _("评审审核—起始日期");
          str = wxT("start_date");
          i_format = 5;
          break;
     case 26:
          str_head = _("评审审核-完成日期");
          str = wxT("finish_date");
          i_format = 5;
          break;
    case 27:
          str_head = _("线下PO确认—负责人");
          str = wxT("name");
          i_format = 1;
          break;
     case 28:
          str_head = _("线下PO确认—起始日期");
          str = wxT("start_date");
          i_format = 5;
          break;
     case 29:
          str_head = _("线下PO确认-完成日期");
          str = wxT("finish_date");
          i_format = 5;
          break;
     case 30:
          str_head = _("PO确认完成—负责人");
          str = wxT("name");
          i_format = 1;
          break;
     case 31:
          str_head = _("PO确认完成—起始日期");
          str = wxT("start_date");
          i_format = 5;
          break;
     case 32:
          str_head = _("PO确认完成-完成日期");
          str = wxT("finish_date");
          i_format = 5;
          break;
     default:
          str_head.Empty();
          str.Empty();
          i_format = 0;
          break;

       }
       array_head.Add(str_head);
       array_str.Add(str);
       array_format.Add(i_format);
    }
}

void wfouttoexcelDialog::export_review(wxString str_file)
{
    datepickerdlg dlg;

    if(dlg.ShowModal()==wxID_CANCEL)
    {
        return;
    }

    bool b_from = dlg.b_from;
    bool b_to = dlg.b_to;

    wxDateTime dt_from,dt_to;

    if(b_from)
        dt_from = dlg.dt_from;

    if(b_to)
        dt_to = dlg.dt_to;

    wxString s_sql_att;

    if(b_from && b_to)
        s_sql_att =  wxT(" create_date <='")+DateToAnsiStr(dt_to)+wxT("' and create_date >='")+DateToAnsiStr(dt_from)+wxT("' ");

    if(b_from && !b_to)
        s_sql_att = wxT(" create_date >='")+DateToAnsiStr(dt_from)+wxT("' ");

    if(!b_from && b_to)
        s_sql_att = wxT(" create_date <='")+DateToAnsiStr(dt_to)+wxT("' ");


    wxString str_sql = wxT("SELECT *,(select name from s_employee where employee_id=operator_id) as name, (select name from s_employee where employee_id = res_cm) as res_cm_name,\
                           (select name from s_employee where employee_id = review_engineer) as review_engineer_name \
                           FROM v_task_out2 WHERE workflow_id='WF0001' and active_status>=0 and unit_status>=0 and is_latest=true  and ")+s_sql_att+wxT(" ORDER BY review_task_id,wbs_no, flow_ser ASC;");
    wxPostgreSQLresult *  res = wxGetApp().app_sql_select(str_sql);


    if(res->Status()!= PGRES_TUPLES_OK)
    {
        res->Clear();
        return;
    }

    int i_count = res->GetRowsNumber();

  //  wxLogMessage(NumToStr(i_count));


    int i_unit=0;

    bool b_continue;
    wxString temp_task_id, temp_wbs;

    workbook wb1;
    worksheet* ws = wb1.sheet("Sheet1");


    int i_flow_ser;
    s_struct * result;
    int i ;
    for(i = 0;i<33;i++)
    {
        ws->label(0,i,array_head.Item(i).ToStdWstring());
    }

    res->MoveFirst();
    result = read_review_res(res);
    b_continue = false;

    gauge_process->SetRange(i_count);
    for(i=0;i<i_count;i++)
    {
        i_flow_ser = result->i_flow_ser;
 //       i_total_flow = result->i_total_flow;

        gauge_process->SetValue(i+1);
        if(temp_wbs!= result->array_value.Item(3))
        {
            i_unit++;
            b_continue = false;
            temp_task_id = result->array_value.Item(0);
            temp_wbs = result->array_value.Item(3);

            label_review_result(i_unit,ws,result);

            wxLogMessage(_("输出")+temp_task_id+_("-")+temp_wbs);
        }


        if(b_continue)
        {
             res->MoveNext();
            delete result;
            result =read_review_res(res);
            continue;
        }

        if(temp_wbs== result->array_value.Item(3))
        {
            if(result->b_active)
                b_continue = true;

            ws->label(i_unit,15+(i_flow_ser-1)*3, result->array_value.Item(15).ToStdWstring());
            ws->label(i_unit,15+(i_flow_ser-1)*3+1, result->array_value.Item(16).ToStdString());
            ws->label(i_unit,15+(i_flow_ser-1)*3+2, result->array_value.Item(17).ToStdString());
        }

        res->MoveNext();
        delete result;
        result = read_review_res(res);
    }

    std::string filename = str_file.ToStdString();
    int error_msg = wb1.Dump(filename);

    if(error_msg == NO_ERRORS)
    {
            wxLogMessage(_("共计输出：")+NumToStr(i_unit));
    }


    res->Clear();
    gauge_process->SetValue(0);
}

s_struct* wfouttoexcelDialog::read_review_res(wxPostgreSQLresult *_res)
{
    s_struct* wf_result = new s_struct;
    wf_result->b_active = _res->GetBool(wxT("is_active"));
    wf_result->i_flow_ser = _res->GetInt(wxT("flow_ser"));
    wf_result->i_total_flow = _res->GetInt(wxT("total_flow"));
    wf_result->s_action_id = _res->GetVal(wxT("action_id"));
    wf_result->s_workflow_id = _res->GetVal(wxT("workflow_id"));
    wxString str;
    int i_status;

    for(int i = 0; i< 18; i++)
    {

        if(i==7)
        {
            str = urgent_level_to_str(_res->GetInt(array_str.Item(i)));
        }
        else if(i==12)
        {
            str =prj_status_to_str(_res->GetInt(array_str.Item(i)));
        }
        else
        {

            switch(array_format.Item(i))
            {
            case 0:
            case 1:
                str = _res->GetVal(array_str.Item(i));
                break;
            case 5:
                str = DateToAnsiStr(_res->GetDateTime(array_str.Item(i)));
                break;
            case 8:
                str = DateToStrFormat(_res->GetDateTime(array_str.Item(i)));
                break;
            default:
                str.Empty();
                break;
            }
        }

        wf_result->array_value.Add(str);
    }

    return wf_result;
}

void wfouttoexcelDialog::label_review_result(int irow,worksheet* _ws, s_struct* s_res)
{

    for(int i = 0; i<15;i++)
    {
        if(array_format.Item(i)== 0)
        {
             _ws->label(irow, i, s_res->array_value.Item(i).ToStdString());
        }

        if(array_format.Item(i)== 1)
        {
            _ws->label(irow,i, s_res->array_value.Item(i).ToStdWstring());
        }

        if(array_format.Item(i) == 5)
        {
            _ws->label(irow,i, s_res->array_value.Item(i).ToStdString());
        }

        if(array_format.Item(i)== 8)
        {
            _ws->label(irow, i, s_res->array_value.Item(i).ToStdString());
        }
    }
}

void wfouttoexcelDialog::OnButton_Review_hisClick(wxCommandEvent& event)
{
}

void wfouttoexcelDialog::OnButton_bom_changeClick(wxCommandEvent& event)
{
        wxString str_file;
    str_file = TextCtrl1->GetValue()+wxT("-bomchange.xls");

    if(str_file.IsEmpty())
        return;

    wxLogMessage(_("导出文件:")+str_file);

    init_bom_change_info_header();

    export_bom_change_info(str_file);
}

void wfouttoexcelDialog::init_bom_change_info_header()
{
    array_head.Clear();
    wxString str_head;
    for(int i=0;i<15;i++)
    {
       switch(i)
       {
     case 0:
          str_head = _("更改单号");
          break;
     case 1:
          str_head = _("内容");
          break;
     case 2:
          str_head = _("申请人");
          break;
     case 3:
          str_head = _("申请日期");
          break;
     case 4:
          str_head = _("发布人");
          break;
     case 5:
          str_head = _("发布日期");
          break;

     case 6:
          str_head = _("PU");
          break;
     case 7:
          str_head = _("PP");
          break;

    case 8:
          str_head = _("IE");
          break;
     case 9:
          str_head = _("QM");
          break;
     case 10:
           str_head = _("LO");
           break;
     case 11:
           str_head  = _("PSM");
           break;
     case 12:
            str_head = _("PM");
            break;
     case 13:
             str_head = _("CO");
             break;
     case 14:
           str_head = _("OTHER");
           break;
     default:
          str_head.Empty();
          break;

       }
       array_head.Add(str_head);
    }
}

void wfouttoexcelDialog::export_bom_change_info(wxString str_file)
{
   wxString str_sql = wxT("SELECT bci_id, bci_content, apply_person,(select name from s_employee where employee_id=apply_person) as apply_name, apply_date, publisher, \
       (select name from s_employee where employee_id=publisher) as publish_name, publish_date, is_active, is_published, is_pu, is_qm, is_ie, is_pp, \
       is_lo, is_cs, is_spu, is_mdm, is_others, is_major, is_psm, is_pm, is_co\
       FROM l_bom_change_inform_table where is_published=true and is_active=true order by bci_id asc;");

    wxPostgreSQLresult * res = wxGetApp().app_sql_select(str_sql);

    if(res->Status()!= PGRES_TUPLES_OK)
    {
        res->Clear();
        return;
    }


    workbook wb1;
    worksheet* ws = wb1.sheet("Sheet1");

    int i_flow_ser;
    s_struct * result;
    int i ;

    for(i = 0;i<15;i++)
    {
        ws->label(0,i,array_head.Item(i).ToStdWstring());
    }

    int i_count = res->GetRowsNumber();

    gauge_process->SetRange(i_count);
    res->MoveFirst();
    wxString str;
    wxDateTime dt_temp;
//操作代码
    for(int i=0; i<i_count; i++)
    {
        gauge_process->SetValue(i+1);
        str  = res->GetVal(wxT("bci_id"));

        ws->label(i+1, 0, str.ToStdString());

        str = res->GetVal(wxT("bci_content"));

        ws->label(i+1, 1, str.ToStdWstring());

        str = res->GetVal(wxT("apply_name"));

        ws->label(i+1, 2, str.ToStdWstring());

        wxDateTime dt_temp= res->GetDate(wxT("apply_date"));

        ws->label(i+1, 3, DateToStrFormat(dt_temp).ToStdString());

        str = res->GetVal(wxT("publish_name"));

         ws->label(i+1, 4, str.ToStdWstring());

        dt_temp = res->GetDate(wxT("publish_date"));
        str = DateToStrFormat(dt_temp);
        ws->label(i+1, 5, str.ToStdString());

        bool b_val = res->GetBool(wxT("is_pu"));
        ws->label(i+1, 6, BoolToY(b_val).ToStdString());

         b_val = res->GetBool(wxT("is_pp"));
         ws->label(i+1, 7, BoolToY(b_val).ToStdString());

         b_val = res->GetBool(wxT("is_ie"));
         ws->label(i+1, 8, BoolToY(b_val).ToStdString());

         b_val = res->GetBool(wxT("is_qm"));
         ws->label(i+1, 9, BoolToY(b_val).ToStdString());

         b_val = res->GetBool(wxT("is_lo"));
         ws->label(i+1, 10, BoolToY(b_val).ToStdString());

         b_val = res->GetBool(wxT("is_psm"));
         ws->label(i+1,11, BoolToY(b_val).ToStdString());

         b_val = res->GetBool(wxT("is_pm"));
         ws->label(i+1, 12, BoolToY(b_val).ToStdString());

         b_val = res->GetBool(wxT("is_co"));
         ws->label(i+1, 13, BoolToY(b_val).ToStdString());

         b_val = res->GetBool(wxT("is_others"));
         ws->label(i+1, 14, BoolToY(b_val).ToStdString());


        res->MoveNext();

    }

//操作代码

    std::string filename = str_file.ToStdString();
    int error_msg = wb1.Dump(filename);

    if(error_msg == NO_ERRORS)
    {
            wxLogMessage(_("共计输出：")+NumToStr(i_count));
    }


    res->Clear();
    gauge_process->SetValue(0);
}
