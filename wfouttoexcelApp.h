/***************************************************************
 * Name:      wfouttoexcelApp.h
 * Purpose:   Defines Application Class
 * Author:     ()
 * Created:   2014-03-27
 * Copyright:  ()
 * License:
 **************************************************************/

#ifndef WFOUTTOEXCELAPP_H
#define WFOUTTOEXCELAPP_H

#include <wx/app.h>
#include "interface/wxPostgreSQL.h"
#include "pugixml.hpp"

typedef struct prog_para
{
    wxString login_user;
    wxString login_role;
    wxString login_user_name;
    wxString plant;
    wxString lang_info;
    bool login_status;
    wxDateTime login_datetime;
    wxString server_host;
    wxString server_port;
    wxString server_dbname;
    wxString local_user;
    wxString local_ip;
    wxString local_pc_name;
} para;

extern para gr_para;

class wfouttoexcelApp : public wxApp
{
    public:
        virtual bool OnInit();

        virtual int OnExit();

        bool read_para();
        wxPostgreSQL* conn;

        bool app_sql_update(wxString &_query);//include insert and update)
        wxPostgreSQLresult * app_sql_select(wxString & _query);
    private:
        wxString conn_str;
};

DECLARE_APP(wfouttoexcelApp)
#endif // WFOUTTOEXCELAPP_H
