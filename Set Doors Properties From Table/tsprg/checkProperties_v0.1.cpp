/*******************************************************************************
***
***     Author: Ivan Matveev
***     email:  matveyev.ivan@gmail.com
***
***     Version: 0.1
***     ��������� ����������� �������� ����� � ���� � ��������� ������� � � ������� �������� � LabPP Automat
***
***     2021
***
*******************************************************************************/
// ���������� ��� � �������������� ���� ������ � � ������� ����� � �����_2.xlsx�

string sExcelGUIDs = "���� � �����_2.xlsx";
int iTableGUIDs; // ����� ������� � GUID��� � ����������. ��������� � ������� Excel ��������� � ������ ������.
int tableRowsNumber;

int main()
{
    int ires;
    int icount;

    //--------------------------------------------------
    // ������� �������
    //--------------------------------------------------

    createTable();
    ts_table(iTableGUIDs, "get_rows_count", tableRowsNumber);
    cout << "\n";









    ac_request_special("load_elements_list_from_selection", 1, "DoorType", 2);

    ac_request("get_loaded_elements_list_count", 1); // ������� ���������� ���������, ����������� � ������ �1
    icount = ac_getnumvalue(); // �������� � ���������� icount ��������� ���������� �������� ��� �����
    cout << "���������� ��������� ������: " << icount << "\n"; // ������ ��������� � ���������� ���������

    if (icount == 0)
    {
        // ������ � ������ ��� ���������, ������� �������� � ���� ��������� � ��������� ���������
        cout << "��� ��������� ��� ���������\n";
        return -1; // ����� ����� �� �����, �� ������ ���� ������������ ������������� ��������, ������ ��������� �� ������� ��, ��� ��������.
    }

    cout << "������ ������ ��������.\n";
    cout << "�������� �� ��������� ������. \n";
    cout << "\n";



    // ..######..##....##..######..##.......########
    // .##....##..##..##..##....##.##.......##......
    // .##.........####...##.......##.......##......
    // .##..........##....##.......##.......######..
    // .##..........##....##.......##.......##......
    // .##....##....##....##....##.##.......##......
    // ..######.....##.....######..########.########




    for (int i = 0; i < icount; i++)
    {
        ac_request("set_current_element_from_list", 1, i); // ���������� ������� ������� �� ������ � 1 � �������� i

        //--------------------------------------------------
        // ��������, � ����� ����� ��������� �������
        //--------------------------------------------------

        cout << "�������� ���� �����.\n";
        string sGUID, 
                curObjFromGUID, 
                curObjToGUID, 
                curDBObjFromNumber, 
                curDBObjToNumber, 
                curDBObjFromCat, 
                curDBObjToCat, 
                curDBObjFromName, 
                curDBObjToName;

        ac_request("get_element_value", "GuidAsText");   // ��������� guid �������� �������� ��� �����
        sGUID = ac_getstrvalue();

        // ��� ���� �������� ���� ������� �����

        curObjFromGUID = getLinkedZoneGUID (sGUID, true); // �������� ���� ����: true ������ ����� �����, false � ���� ���� �����
        curObjToGUID = getLinkedZoneGUID (sGUID, false);  // �������� ���� ����: true ������ ����� �����, false � ���� ���� �����

        // �������� ��� � ������ ��������� ��������� ���

        if (curObjToGUID == "") { cout << "   ���� ������ �� ���������.\n"; }
        if (curObjFromGUID == "") { cout << "   ���� ����� �� ���������.\n"; }


        // 
        // 
        // 
        // 
        // 
        // GET CURRENT OBJECT PARAMS FROM LABBPP DATABASE
        // 
        // 
        // 
        // 
        // 

        curDBObjFromNumber = "";
        curDBObjFromName = "";
        curDBObjFromCat = "";

        curDBObjToNumber = "";
        curDBObjToName = "";
        curDBObjToCat = "";


        if (curObjToGUID != "")
        {
            ac_request("set_element_by_guidstr_as_current", curObjToGUID); // here is the set current element, beware 
            ac_request("get_element_value", "ZoneNumber"); // ����� ����. ����������, �� ������������.
            curDBObjToNumber = ac_getstrvalue();
            ac_request("get_element_value", "ZoneName");
            curDBObjToName = ac_getstrvalue();
            ac_request("get_element_value", "ZoneCatCode");
            curDBObjToCat = ac_getstrvalue();
            cout << "    ZONE TO GUID = " << curObjToGUID << ".\n";
            cout << "    ZONE TO NAME = " << curDBObjToName << ".\n";
        }

        if (curObjFromGUID != "")
        {
            ac_request("set_element_by_guidstr_as_current", curObjFromGUID); // here is the set current element, beware 
            ac_request("get_element_value", "ZoneNumber");

            curDBObjFromNumber = ac_getstrvalue();
            ac_request("get_element_value", "ZoneName");
            
            curDBObjFromName = ac_getstrvalue();
            ac_request("get_element_value", "ZoneCatCode");
            curDBObjFromCat = ac_getstrvalue();

            cout << "    ZONE FROM GUID = " << curObjFromGUID << ".\n";
            cout << "    ZONE FROM NAME = " << curDBObjFromName << ".\n";
        }

        // ZONE_FROM_CATEGORY  ZONE_FROM_NAME  ZONE_TO_CATEGORY    ZONE_TO_NAME    DOOR_CATEGORY

        // 
        // 
        // 
        // 
        // 
        // 
        // 
        // GET OBJECT'S OWN SET PROPERTIES
        // 
        // 
        // 
        // 
        // 
        // 
        // 

        //
        // �������� ����� �� ��������������:
        //
        // DOOR_CATEGORY � ��������� �����
        //
        // ZONE_FROM_NAME � ��� ����, ������ ����� �����
        // ZONE_FROM_CATEGORY � ��������� ����, ������ ����� �����
        // ZONE_FROM_GUID � GUID ����, ������ ����� �����
        //
        // ZONE_TO_NAME � ��� ����, ���� ����� �����
        // ZONE_TO_CATEGORY � ��������� ����, ���� ����� �����
        // ZONE_TO_GUID � GUID ����, ���� ����� �����
        //
        //---------------------------------------

        string  curObjPropFromName, 
                curObjPropFromCat,
                curObjPropFromGUID,
                curObjPropToName, 
                curObjPropToCat,
                curObjPropToGUID,
                curDoorCategory;

        curObjPropFromName = "";
        curObjPropFromCat = "";
        curObjPropFromGUID = "";
        curObjPropToName = "";
        curObjPropToCat = "";
        curObjPropToGUID = "";
        curDoorCategory = "";


        ac_request("set_current_element_from_list", 1, i);
        ires = ac_request("elem_user_property", "get", "DOOR_CATEGORY");
            curDoorCategory = ac_getstrvalue();

        ires = ac_request("elem_user_property", "get", "ZONE_FROM_NAME");
            curObjPropFromName = ac_getstrvalue();
        ires = ac_request("elem_user_property", "get", "ZONE_FROM_CATEGORY");
            curObjPropFromCat = ac_getstrvalue();
        ires = ac_request("elem_user_property", "get", "ZONE_FROM_GUID");
            curObjPropFromGUID = ac_getstrvalue();

        ires = ac_request("elem_user_property", "get", "ZONE_TO_NAME");
            curObjPropToName = ac_getstrvalue();
        ires = ac_request("elem_user_property", "get", "ZONE_TO_CATEGORY");
            curObjPropToCat = ac_getstrvalue();
        ires = ac_request("elem_user_property", "get", "ZONE_TO_GUID");
            curObjPropToGUID = ac_getstrvalue();
        


        // �������� �� ������� �������� � 

        if (curObjPropFromName == " ") { // ���������� �� �������: ��� ������ ���� � ��������
            curObjPropFromName = "";
        }
        if (curObjPropFromCat == " ") {
            curObjPropFromCat = "";
        }
        if (curObjPropToName == " ") {
            curObjPropToName = "";
        }
        if (curObjPropToCat == " ") {
            curObjPropToCat = "";
        }
        if (curDoorCategory == " ") {
            curDoorCategory = "";
        }


        // 
        // 
        // 
        // 
        // CHECK IF OBJECT'S PROPERTIES AND ITS LABPP ZONES DIFFER
        // 
        // 
        // 
        // 

        int e1, e2, e3, e4, e5;

        bool whaaaaa = false;

        if (curObjPropFromName != curDBObjFromName) {
            cout << "! �������� �������: �" << curObjPropFromName << "� != �" << curDBObjFromName;
            cout << "\n";
            whaaaaa = true;
        } else {
            cout << "   OK, �������� �������: �" << curObjPropFromName << "� == �" << curDBObjFromName;
            cout << "\n";
        }
        if (curObjPropFromCat != curDBObjFromCat) {
            cout << "! �������� �������: �" << curObjPropFromCat << "� != �" << curDBObjFromCat;
            cout << "\n";
            whaaaaa = true;
        } else {
            cout << "   OK, �������� �������: �" << curObjPropFromCat << "� == �" << curDBObjFromCat;
            cout << "\n";
        }
        if (curObjPropToName != curDBObjToName) {
            cout << "! �������� �������: �" << curObjPropToName << "� != �" << curDBObjToName << "�";
            cout << "\n";
            whaaaaa = true;
        } else {
            cout << "   OK, �������� �������: �" << curObjPropToName << "� == �" << curDBObjToName;
            cout << "\n";
        }
        if (curObjPropToCat != curDBObjToCat) {
            cout << "! �������� �������: �" << curObjPropToCat << "� != �" << curDBObjToCat;
            cout << "\n";
            whaaaaa = true;
        } else {
            cout << "   OK, �������� �������: �" << curObjPropToCat << "� == �"  << curDBObjToCat;
            cout << "\n";
        }





        // 
        // 
        // 
        // 
        // 
        // 
        // 
        // 
        // 
        // 
        // 
        // 
        // 
        // 
        // 
        // 
        // 
        // 
        // 
        // 
        // 


        // ..######..########....###....########...######..##.....##
        // .##....##.##.........##.##...##.....##.##....##.##.....##
        // .##.......##........##...##..##.....##.##.......##.....##
        // ..######..######...##.....##.########..##.......#########
        // .......##.##.......#########.##...##...##.......##.....##
        // .##....##.##.......##.....##.##....##..##....##.##.....##
        // ..######..########.##.....##.##.....##..######..##.....##

        //--------------------------------------------------
        // �����!
        //--------------------------------------------------
        // ������ � �������, � ������������� �������� ���� ��� � ���������� (�����) �����
        //--------------------------------------------------

        // int irow = ts_table(iTableGUIDs, "search", 0, currentObjectZoneCombination);

        // ZONE_FROM_NAME � ��� ����, ������ ����� �����
        // ZONE_FROM_CATEGORY � ��������� ����, ������ ����� �����
        // ZONE_FROM_GUID � GUID ����, ������ ����� �����

        // ZONE_TO_NAME � ��� ����, ���� ����� �����
        // ZONE_TO_CATEGORY � ��������� ����, ���� ����� �����
        // ZONE_TO_GUID � GUID ����, ���� ����� �����

        bool doorWithoutCategory = true;
        string sIDdoor, curObjFromNumber, curObjToNumber, curObjFromCat, curObjToCat, curObjFromName, curObjToName;

        string  curTableFromCat, curTableToCat,
                curTableFromName, curTableToName;

        // ��������� �������:
        // [ ZONE_FROM_CATEGORY | ZONE_FROM_NAME | ZONE_TO_CATEGORY | ZONE_TO_NAME | DOOR_CATEGORY ]
        // [ 7 | zone 1 | 3 | zone 5 | 015 ]

        
        ac_request("set_current_element_from_list", 1, i);
        cout << "   ���������� ���������...\n";

        //
        // �������� ����� �� ��������������:
        //
        // DOOR_CATEGORY � ��������� �����
        //
        // ZONE_FROM_NAME � ��� ����, ������ ����� �����
        // ZONE_FROM_CATEGORY � ��������� ����, ������ ����� �����
        // ZONE_FROM_GUID � GUID ����, ������ ����� �����
        //
        // ZONE_TO_NAME � ��� ����, ���� ����� �����
        // ZONE_TO_CATEGORY � ��������� ����, ���� ����� �����
        // ZONE_TO_GUID � GUID ����, ���� ����� �����
        //
        //---------------------------------------

        

        if (curObjFromGUID != "")
        {
            ires = ac_request("elem_user_property", "set", "ZONE_FROM_NAME", curObjFromName);
            // cout << "       ��� ���� > " << ires << "\n";
            ires = ac_request("elem_user_property", "set", "ZONE_FROM_CATEGORY", curObjFromCat);
            // cout << "       ��������� ���� >" << ires << "\n";
            ires = ac_request("elem_user_property", "set", "ZONE_FROM_GUID", curObjFromGUID);
            // cout << "       GUID ���� ������ >" << ires << "\n";
        } else {
            ires = ac_request("elem_user_property", "set", "ZONE_FROM_NAME", "");
            ires = ac_request("elem_user_property", "set", "ZONE_FROM_CATEGORY", "");
            ires = ac_request("elem_user_property", "set", "ZONE_FROM_GUID", "");
        }

        if (curObjToGUID != "")
        {
            ires = ac_request("elem_user_property", "set", "ZONE_TO_NAME", curObjToName);
            // cout << "       ���� < " << ires << "\n";
            ires = ac_request("elem_user_property", "set", "ZONE_TO_CATEGORY", curObjToCat);
            // cout << "       ��������� ���� < " << ires << "\n";
            ires = ac_request("elem_user_property", "set", "ZONE_TO_GUID", curObjToGUID);
            // cout << "       GUID ���� < " << ires << "\n";
        } else {
            ires = ac_request("elem_user_property", "set", "ZONE_TO_NAME", "");
            ires = ac_request("elem_user_property", "set", "ZONE_TO_CATEGORY", "");
            ires = ac_request("elem_user_property", "set", "ZONE_TO_GUID", "");
        }


        for (int row = 0; row < tableRowsNumber; row++) // cycle through all table rows
        {
            cout << "       Row: " << row+1 << " / " << tableRowsNumber << "\n";
            ts_table(iTableGUIDs, "select_row", row); // set the row

            curTableFromCat = ""; // ���������� ����������, ����� �� ������ � ��������� ��������� � ���������
            curTableFromName = "";
            curTableToCat = "";
            curTableToName = "";
            curDoorCategory = "";

            ts_table(iTableGUIDs, "get_value_of", 0, curTableFromCat);  // get current zone_from cat. in this row
            ts_table(iTableGUIDs, "get_value_of", 1, curTableFromName); // get current zone_from name in this row
            ts_table(iTableGUIDs, "get_value_of", 2, curTableToCat);    // get current zone_from name in this row
            ts_table(iTableGUIDs, "get_value_of", 3, curTableToName);   // get current zone_from name in this row
            ts_table(iTableGUIDs, "get_value_of", 4, curDoorCategory);  // get current zone_from name in this row


            curTableFromCat = tolower(curTableFromCat);
            curTableFromName = tolower(curTableFromName);
            curTableToCat = tolower(curTableToCat);
            curTableToName = tolower(curTableToName);
            curDoorCategory = tolower(curDoorCategory);

            curObjFromCat = tolower(curObjFromCat);
            curObjFromName = tolower(curObjFromName);
            curObjToCat = tolower(curObjToCat);
            curObjToName = tolower(curObjToName);


        
            // �������������� ����� �������� �� �������
            // cout << "!!!!! Current zone from category: " <<  curTableFromCat << "\n";
            // cout << "!!!!! Current zone from name: " <<  curTableFromName << "\n";
            // cout << "!!!!! Current zone to category: " <<  curTableToCat << "\n";
            // cout << "!!!!! Current zone to name: " <<  curTableToName << "\n";
            // cout << "!!!!! Current door category: " <<  curDoorCategory << "\n";

            // ..######..########.########
            // .##....##.##..........##...
            // .##.......##..........##...
            // ..######..######......##...
            // .......##.##..........##...
            // .##....##.##..........##...
            // ..######..########....##...

            // �������
            // curTableFromName, curTableFromCat, curTableToName, curTableToCat � ��� ��������� �� ������� ������ �������
            // ������
            // curObjFromName, curObjFromCat, curObjToName, curObjToCat � ��� ��������� �������� �������

            if (curTableFromCat == "*") { // ���������� �� �������: ��� ������ ���� � ��������
                curTableFromCat = "";
            }
            if (curTableFromName == "*") {
                curTableFromName = "";
            }
            if (curTableToCat == "*") {
                curTableToCat = "";
            }
            if (curTableToName == "*") {
                curTableToName = "";
            }

            string currentObjectFromToCatName, currentTableFromToCatName;


            currentObjectFromToCatName = "_|_" + curTableFromCat + "_|_" + curTableFromName + "_|_" + curTableToCat + "_|_" + curTableToName + "_|_";

            currentTableFromToCatName = "_|_" + curObjFromCat + "_|_" + curObjFromName + "_|_" + curObjToCat + "_|_" + curObjToName + "_|_";


            // ���� ��� ��������� ���������, �� �������� � ����� ��������� �����

            
            //---------------------------------------
            //
            // ������ �������� �����
            //
            //---------------------------------------



            //  ��� ����� � ������!!!
            if (currentObjectFromToCatName == currentTableFromToCatName) {
                ires = ac_request("elem_user_property", "set", "DOOR_CATEGORY", curDoorCategory);
                cout << "   ���������� ��������� �����: " << ires << "\n";
                doorWithoutCategory = false;
            }

        } // end of zones for loop







        /* ����������������� ����� ����������� �������!!

        // ��� ���� ��������� �� 

        if (curDoorCategoryFromTable != curDoorCategory) {
            cout << "! �������� �������: �" << curObjPropToCat << "� != �" << curDBObjToCat;
            cout << "\n";
            whaaaaa = true;
        } else {
            cout << "   OK, �������� �������: �" << curObjPropToCat << "� == �"  << curDBObjToCat;
            cout << "\n";
        }
    */

        // �������� �� ������������ ���� �����
        // 
        // 
        // 
        // 
        // ���������� ������, ����� ��� �����. ������, 













        if (whaaaaa == true) {

            // ��� ���� ������� � ������ �������� ����� ������ ��� ���, 
            // ����� ���������� ������� �������� ����� �����������
            // ������� � 
            // ...

            int r0, ires;
            bool the_res = true;

            r0 = ac_request_special("get_element_value", "UP", "Wrong Category/Type");
            // cout << r0 << "\n";

            if (r0 == -1222) {  //-1222 � ������ �����-��
                cout << "���� �� ������� �� �������, ���� �������.\n";
                ires = ac_request("elem_user_property", "create", "Wrong Category/Type", " ", "Boolean", "OPENINGS");
                if (ires == 0) { cout << "������� �������� Wrong Category/Type.\n"; } else { cout << "�� ����� ������� �������� Wrong Category/Type.\n"; break();}
            }


            cout << "   ������ � �������: " << sGUID << "\n";
            ires = ac_request("elem_user_property", "set", "Wrong Category/Type", 1); // ��������� ������ ��� �������� � ��������, ������ ��� ����������� ������� � AC



        // <!-- �� ����� ��� ������, ���������� ����������
        // �������� ������ � ������� � ������ 2
        
/*
            [16:33, 09/11/2021] ���� �����: ������ ����!)
������
[16:36, 09/11/2021] ���� �����: ac_request("store_current_element_to_list",2,-1);
    ac_request("select_elements_from_list",2,1);
    ac_request("Automate","ZoomToElements",2);
[16:37, 09/11/2021] ���� �����: ac_request("clear_list",2);
[16:37, 09/11/2021] ���� �����: ������� ������� ������ ����� �����-��
[16:37, 09/11/2021] ���� �����: ����� - �2
[16:38, 09/11/2021] ���� �����: ����� ����� ���������� �������� � �����, ��������, ������ "store_current_element_to_list"
[16:38, 09/11/2021] ���� �����: �.�. ������� ������� ���� ���������� - ���������� ��� � ������ 2
[16:39, 09/11/2021] ���� �����: ��� -1 ������ � ����� ������
[16:39, 09/11/2021] ���� �����: ����� ������ select_elements_from_list
[16:39, 09/11/2021] ���� �����: ��� ����� ������ 2 � 1/0 - ������� ��� ��� ������� ���������
[16:39, 09/11/2021] ���� �����: ��� ������� ����� ZoomToElements �������
[16:40, 09/11/2021] ���� �����: ���� � 3d ���� - �������� �����
[16:40, 09/11/2021] ���� �����: � � 2d ���� ���� ������� ���� ����������
[16:40, 09/11/2021] ���� �����: ac_request("Environment","Story_GoTo",storeindex);
[16:41, 09/11/2021] ���� �����: storeindex - ������ �����
[16:41, 09/11/2021] ���� �����: ��� � ����� ������ ������, �������.
����� �� ���� ���� �������������.
[16:42, 09/11/2021] ���� �����: � ��� ����� ������ ���������, ��������� � �.�.
[16:43, 09/11/2021] ���� �����: ����� ������������ ������� ts_table � guid-��� � �������� ��� � ������ 2

*/
// �� ����� ��� ������, ���������� ���������� -->


        } else {
            ires = ac_request("elem_user_property", "set", "Wrong Category/Type", 0); // �������� ����
        }
        cout << "\n";







    }

    // �������� ������� �� ������ 2
    // 
    // 
    // 
    //  �� ����� ��� ������, ���������� ����������
    // 
    // 


    deleteTable();

    cout << "\n";
    cout << "��������� ���������� �������.\n";
}










// .########.##.....##.##....##..######..########.####..#######..##....##..######.
// .##.......##.....##.###...##.##....##....##.....##..##.....##.###...##.##....##
// .##.......##.....##.####..##.##..........##.....##..##.....##.####..##.##......
// .######...##.....##.##.##.##.##..........##.....##..##.....##.##.##.##..######.
// .##.......##.....##.##..####.##..........##.....##..##.....##.##..####.......##
// .##.......##.....##.##...###.##....##....##.....##..##.....##.##...###.##....##
// .##........#######..##....##..######.....##....####..#######..##....##..######.





int LoadExcel()
{
    // ������ IDispatch GetIDsOfNames....
    // [17:56, 27/10/2021] ���� �����: � ���� �������
    // [17:56, 27/10/2021] ���� �����: 1 - ���� Excel ��������� � ������ �������������� ������
    // [17:57, 27/10/2021] ���� �����: 2 - ������ ��� �����-�� Excel � ���� ��������
    // [17:58, 27/10/2021] ���� �����: ������� - ��������� ��������� ����� � �������� Excel � ���������
    // [17:58, 27/10/2021] ���� �����: ��� ������� - Excel ��������� �� ����� ��������������
    // [17:58, 27/10/2021] ���� �����: ���������� ���������

    int i, tsalert, res;
    string sguid, str;

    cout << "��������� ������ �� Excel\n";

    //  runtimecontrol("workline", "setpos", 0);

    // ����� �������� ������ �������� �������: ������ ���������� Excel ��� ������.
    // � ������ ������� ��������� Excel ������ ���� �������.
    // �� ��������� �������� ���������� ������� �������� Excel.


    res = excel_attach();

    if (res != 0)
    {
        tsalert(-1, "������ �� ����� ����������", "��� ����������� � excel", "");
        return -1;
    }


    res = excel_request("workbook_select", sExcelGUIDs);

    if (res != 0)
    {
        tsalert(-1, "������ �� ����� ����������", "�� ���������� ������������� � ���� excel", sExcelGUIDs);
        excel_detach();
        return -1;
    }

    ts_table(iTableGUIDs, "import_columns_from_excel", "A", 1, -1);  // � ������ ������ ������� �, �� ������ ������ ������
    ts_table(iTableGUIDs, "import_from_excel", "A", 2, -1, 0, 1);  // �� ������ ������ ������� �, -1 � � ������� ������, ... , �������� ������� ����� �����������

    excel_detach();

    cout << "�������� ���������.\n";

    //ts_table(iTableGUIDs, "print_to_str", str); // ��������� ��� ������� � ��������� ���������� ��� ��������
    //cout << "���������� ������� guids ->" << str << "\n"; // ������� � ���� ���������

    return 0;
}


string getLinkedZoneGUID(string sDoorGuid, bool bRoomFrom) // �������� ���� ����: ���� true � ������ ����� �����, ���� false � ���� ���� �����
{
    string sGUID_linked_zone_guid = "";

    int flag1, flag2;
    if (bRoomFrom) { flag1 = 1024; flag2 = 2048; } else { flag1 = 4096; flag2 = 8192; }

    int iGuid;
    object("create", "ts_guid", iGuid);
    ts_guid(iGuid, "ConvertFromString", sDoorGuid);
    int iTableRoomsTmp;
    object("create", "ts_table", iTableRoomsTmp);

    ac_request_special("linkingElems", "getLinkedElemsByFlags", iGuid, flag2, iTableRoomsTmp);
    int icount;
    ts_table(iTableRoomsTmp, "get_rows_count", icount);
    if (icount > 0)
    {
        ts_table(iTableRoomsTmp, "select_row", 0);
        ts_table(iTableRoomsTmp, "get_value_of", "sGUID", sGUID_linked_zone_guid);
    }
    object("delete", iGuid);
    object("delete", iTableRoomsTmp);
    return sGUID_linked_zone_guid;
}

int createTable() {
    string sguid;   // ���������� ��� guid
    string svalue;  // ��� ���������
    int ivalue;     // ��� ��������
    int jcount;
    int icount;
    int ires;
    object("create", "ts_table", iTableGUIDs); // ������� ������ ts_table � �������� ��� ����� ��� ������ � ���
    ires = LoadExcel(); // ��������� ������ �� ������ (��. ����������� ������� ����)
    if (ires != 0) {
        cout << "������ � ���� �������� �� Excel, ��������� ���������";
        return -1;
    }
    cout << "������� �������. \n";
    return 0;
}

int deleteTable() {
    object("delete", iTableGUIDs);  // ������� ������
    cout << "�������� ������ �� �������. \n";
    return 0;
}