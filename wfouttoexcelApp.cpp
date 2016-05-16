/***************************************************************
 * Name:      wfouttoexcelApp.cpp
 * Purpose:   Code for Application Class
 * Author:     ()
 * Created:   2014-03-27
 * Copyright:  ()
 * License:
 **************************************************************/

#include "wfouttoexcelApp.h"
#include "connect_para_dlg.h"
#include "security/base64.h"
#include <wx/socket.h>

//(*AppHeaders
#include "wfouttoexcelMain.h"
#include <wx/image.h>
//*)

IMPLEMENT_APP(wfouttoexcelApp);

para gr_para;
bool wfouttoexcelApp::OnInit()
{

    if(!read_para()) return false;


    bool b_return;

    do
    {

        conn = new wxPostgreSQL(conn_str, wxConvUTF8);

//   int s_ver;
        if(conn->Status()==CONNECTION_OK)
        {
//        wxMessageBox(wxString(conn->GetHost()));
//      s_ver = conn->ServerVersion();
            b_return= true;
        }
        else
        {

            connect_para_dlg dlg;

            if(dlg.ShowModal() == wxID_OK)
            {
                if(conn)
                  delete conn;
                read_para();
                conn = new wxPostgreSQL(conn_str, wxConvUTF8);

                //   int s_ver;
                if(conn->Status()==CONNECTION_OK)
                {
//        wxMessageBox(wxString(conn->GetHost()));
//      s_ver = conn->ServerVersion();
                    b_return= true;
                }else b_return=false;
            }else
            {
              wxMessageBox(conn->ErrorMessage(),"Connect error:");
               return false;
            }

        }
    }
    while(!b_return);

    wxIPV4address addr;
    gr_para.local_pc_name = wxGetFullHostName();
    gr_para.local_user = wxGetUserId();
    gr_para.login_status = false;
    addr.Hostname(gr_para.local_pc_name);
    gr_para.local_ip = addr.IPAddress();
    gr_para.lang_info = wxT("zh_CN");

    //(*AppInitialize
    bool wxsOK = true;
    wxInitAllImageHandlers();
    if ( wxsOK )
    {
    	wfouttoexcelDialog Dlg(0);
    	SetTopWindow(&Dlg);
    	Dlg.ShowModal();
    	wxsOK = false;
    }
    //*)
    return wxsOK;

}

int wfouttoexcelApp::OnExit()
{
    if(conn)
    {
        conn->Finish();
            delete conn;
    }

    return 0;
}


bool wfouttoexcelApp::app_sql_update(wxString &_query)
{
    return conn->ExecuteVoid(_query);
}

wxPostgreSQLresult * wfouttoexcelApp::app_sql_select(wxString & _query)
{
    return conn->Execute(_query);
}


bool wfouttoexcelApp::read_para()//read the parameter from the xml file
{

    wxString ldbusr,ldbpwd;
    pugi::xml_document ldoc;
    pugi::xml_node lnode;
    pugi::xml_parse_result result;

    result = ldoc.load_file("config.xml");
    if(result)
    {
        lnode =ldoc.child("connect");
        gr_para.server_host = wxString(lnode.child_value("host"));
        gr_para.server_port = wxString(lnode.child_value("port"));
        gr_para.server_dbname = wxString(lnode.child_value("dbname"));
        ldbusr = wxString(lnode.child_value("dbuser"));
        ldbpwd = wxString(lnode.child_value("dbpwd"));
        ldbpwd = Base64::Decode(ldbpwd);
//       ldbpwd = wxBase64Encode(ldbpwd,ldbpwd.Length());
//       wxMessageBox(ldbpwd);
        conn_str = wxT("host=")+gr_para.server_host +wxT(" port=")+gr_para.server_port+ wxT(" dbname=")+gr_para.server_dbname + wxT(" user=")+ldbusr + wxT(" password=")+ldbpwd;

        return true;
    }
    else return false;

}
