/***************************************************************
 * Name:      wfouttoexcelMain.h
 * Purpose:   Defines Application Frame
 * Author:     ()
 * Created:   2014-03-27
 * Copyright:  ()
 * License:
 **************************************************************/

#ifndef WFOUTTOEXCELMAIN_H
#define WFOUTTOEXCELMAIN_H

//(*Headers(wfouttoexcelDialog)
#include <wx/sizer.h>
#include <wx/textctrl.h>
#include <wx/button.h>
#include <wx/dialog.h>
#include <wx/gauge.h>
//*)

#include "interface/wxPostgreSQL.h"
#include "xlslib.h"
#include <wx/tokenzr.h>


using namespace xlslib_core;
#define RECORDCOUNT                65536

int CompareWBSString(const wxString& first, const wxString& second);
struct s_struct
{
    ~s_struct()
    {
        array_value.Clear();
    }
    int i_flow_ser;
    int i_total_flow;
    bool b_active;
    wxString s_action_id;
    wxString s_workflow_id;
    wxArrayString array_value;
};

wxString nstd_level_to_str(int i_status);
wxString urgent_level_to_str(int i_level);

wxString combine_str(wxArrayString &a_str ,wxString serial_str, wxString individual_str);

class wfouttoexcelDialog: public wxDialog
{
    public:

        wfouttoexcelDialog(wxWindow* parent,wxWindowID id = -1);
        virtual ~wfouttoexcelDialog();

    private:

        //(*Handlers(wfouttoexcelDialog)
        void OnQuit(wxCommandEvent& event);
        void OnAbout(wxCommandEvent& event);
        void OnButton2Click(wxCommandEvent& event);
        void OnButton3Click(wxCommandEvent& event);
        void OnButton1Click(wxCommandEvent& event);
        void OnButton4Click(wxCommandEvent& event);
        void OnButton5Click(wxCommandEvent& event);
        void OnButton6Click(wxCommandEvent& event);
        void Onbutton_workClick(wxCommandEvent& event);
        void OnButton7Click(wxCommandEvent& event);
        void OnButton_ReasonClick(wxCommandEvent& event);
        void OnButton_ConfClick(wxCommandEvent& event);
        void OnButton_Conf_NstdClick(wxCommandEvent& event);
        void OnButton_ReviewClick(wxCommandEvent& event);
        void OnButton_Review_hisClick(wxCommandEvent& event);
        void OnButton_bom_changeClick(wxCommandEvent& event);
        //*)

        //(*Identifiers(wfouttoexcelDialog)
        static const long ID_TEXTCTRL1;
        static const long ID_BUTTON1;
        static const long ID_BUTTON2;
        static const long ID_BUTTON4;
        static const long ID_BUTTON5;
        static const long ID_BUTTON8;
        static const long ID_BUTTON_CONF;
        static const long ID_BUTTON6;
        static const long ID_BUTTON_CONF_NSTD;
        static const long ID_BUTTON_NSTD_REASON;
        static const long ID_BUTTON7;
        static const long ID_BUTTON3;
        static const long ID_BUTTON_REVIEW;
        static const long ID_BUTTON_REVIEW_RESTART_HISTORY;
        static const long ID_BUTTON_BOM_CHANGE;
        static const long ID_GAUGE_STEP;
        static const long ID_TEXTCTRL2;
        //*)

        //(*Declarations(wfouttoexcelDialog)
        wxButton* Button4;
        wxButton* Button1;
        wxGauge* gauge_process;
        wxButton* Button_Conf;
        wxButton* Button2;
        wxButton* Button_Reason;
        wxButton* Button6;
        wxButton* Button5;
        wxButton* Button_Review;
        wxButton* Button3;
        wxButton* Button7;
        wxButton* Button_Review_his;
        wxButton* Button_bom_change;
        wxTextCtrl* TextCtrl1;
        wxTextCtrl* tc_log;
        wxButton* button_work;
        wxButton* Button_Conf_Nstd;
        //*)

        wxString m_file;
        wxArrayString array_head, array_str;
        wxArrayString array_head1, array_str1,array_head2, array_str2;
        wxArrayInt array_format, array_format1, array_format2;
        void export_excel(wxString str_sql, wxString str_file);


        void init_array_head();
        s_struct* read_res(wxPostgreSQLresult *_res);
        void label_result(int irow, worksheet* _ws, s_struct* s_res);
//非标部分
        void init_nonstd_array_head();
        void export_nonstd_excel(wxString str_file);

//非标分类统计
        void init_nonstd_catalog_head();
        void export_nonstd_catalog_excel(wxString str_file);

//日常工作任务导出
        void init_daily_work_head();
        void export_daily_work(wxString str_file);

//导出重排产项目
        void init_restart_project_head();
        void export_restart_porject(wxString str_file);

        void label_restart_info(int i_row, wxString str, worksheet* ws);
        void label_newest_restart_info(int i_row, wxString str, worksheet* ws);

//导出配置任务评估表
        void init_conf_evaluation();
        void export_conf_evaluation(wxString str_file);

        wxString get_contract_id(wxString s_project_id);

        wxString prj_status_to_str(int i_prj);
        int prj_str_to_status(wxString str_status);
        wxString nstd_status_to_str(int i_status);


//导出非标处理表
    void init_conf_app_deal();
    void export_conf_app_deal(wxString str_file);

    s_struct* read_conf_nstd_app_res(wxPostgreSQLresult *_res);
    void label_conf_nstd_app_result(int irow,worksheet* _ws, s_struct* s_res);

//导出评审记录
     void init_review_header();
     void export_review(wxString str_file);

    s_struct* read_review_res(wxPostgreSQLresult *_res);
    void label_review_result(int irow,worksheet* _ws, s_struct* s_res);

//导出评审重启历史记录
     void init_review_history_header();
     void export_review_history(wxString str_file);

//导出BOM 更改通知单
      void init_bom_change_info_header();
      void export_bom_change_info(wxString str_file);


        DECLARE_EVENT_TABLE()
};

#endif // WFOUTTOEXCELMAIN_H
