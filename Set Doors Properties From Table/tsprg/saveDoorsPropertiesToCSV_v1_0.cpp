/*******************************************************************************
***
***     Author: Ivan Matveev
***     email:  matveyev.ivan@gmail.com
***
***     Version: 1.0
***     ��������� �������� ���� ��������� ������ � ������� CSV
***
***     2021
***
*******************************************************************************/
// ���������� ��� � �������������� ���� ������ � � ������� ����� � �����.xlsx�

string filepath = "D:\\table.csv";
int iFileDescr;

int main()
{
    int ires;
    int icount;


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





    // .########....###....########..##.......########
    // ....##......##.##...##.....##.##.......##......
    // ....##.....##...##..##.....##.##.......##......
    // ....##....##.....##.########..##.......######..
    // ....##....#########.##.....##.##.......##......
    // ....##....##.....##.##.....##.##.......##......
    // ....##....##.....##.########..########.########




    // ������� �������
    int TableDescr1;
    object("create", "ts_table", TableDescr1);

    ts_table(TableDescr1, "add_column", 0, "string", "name");
    ts_table(TableDescr1, "set_first_key", 0);




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
        string sText, sGUID, curObjFromGUID, curObjToGUID, curObjFromNumber, curObjToNumber, curObjFromCat, curObjToCat, curObjFromName, curObjToName;

        ac_request("get_element_value", "GuidAsText");   // ��������� guid �������� �������� ��� �����
        sGUID = ac_getstrvalue();

        // ��� ���� �������� ���� ������� �����

        curObjFromGUID = getLinkedZoneGUID (sGUID, true); // �������� ���� ����: true ������ ����� �����, false � ���� ���� �����
        curObjToGUID = getLinkedZoneGUID (sGUID, false);  // �������� ���� ����: true ������ ����� �����, false � ���� ���� �����

        // �������� ��� � ������ ��������� ��������� ���

        if (curObjToGUID == "") { cout << "   ���� ������ �� ���������.\n"; }
        if (curObjFromGUID == "") { cout << "   ���� ����� �� ���������.\n"; }


        curObjFromNumber = "";
        curObjFromName = "";
        curObjFromCat = "";

        curObjToNumber = "";
        curObjToName = "";
        curObjToCat = "";


        if (curObjToGUID != "")
        {
            ac_request("set_element_by_guidstr_as_current", curObjToGUID);
            ac_request("get_element_value", "ZoneNumber");
            curObjToNumber = ac_getstrvalue();
            ac_request("get_element_value", "ZoneName");
            curObjToName = ac_getstrvalue();
            ac_request("get_element_value", "ZoneCatCode");
            curObjToCat = ac_getstrvalue();
            cout << "    ZONE TO GUID =" << curObjToGUID << ".\n";
            cout << "    ZONE TO NAME=" << curObjToName << ".\n";
        }

        if (curObjFromGUID != "")
        {
            ac_request("set_element_by_guidstr_as_current", curObjFromGUID);
            ac_request("get_element_value", "ZoneNumber");
            curObjFromNumber = ac_getstrvalue();
            ac_request("get_element_value", "ZoneName");
            curObjFromName = ac_getstrvalue();
            ac_request("get_element_value", "ZoneCatCode");
            curObjFromCat = ac_getstrvalue();
            cout << "    ZONE FROM GUID=" << curObjFromGUID << ".\n";
            cout << "    ZONE FROM NAME=" << curObjFromName << ".\n";
        }

        // ZONE_FROM_CATEGORY  ZONE_FROM_NAME  ZONE_TO_CATEGORY    ZONE_TO_NAME    DOOR_CATEGORY


        // ���� ������ ��������, ������ �� �������
        if (curObjFromNumber == "") { curObjFromNumber = "*";}
        if (curObjFromName == "") { curObjFromName = "*";}
        if (curObjFromCat == "") { curObjFromCat = "*";}
        if (curObjToNumber == "") { curObjToNumber = "*";}
        if (curObjToName == "") { curObjToName = "*";}
        if (curObjToCat == "") { curObjToCat = "*";}

        sText = "\"" + curObjFromCat + "\";\"" + curObjFromName + "\";\"" + curObjToCat + "\";\"" + curObjToName + "\"";

        // add sText to index (set-like table column) here

        ts_table(TableDescr1, "add_row", 0, sText);
    }



    // .########.####.##.......########
    // .##........##..##.......##......
    // .##........##..##.......##......
    // .######....##..##.......######..
    // .##........##..##.......##......
    // .##........##..##.......##......
    // .##.......####.########.########


    // ��� ���� ������� ���� � �������� � ���� ��, ��� ���� ��������


    object("create", "ts_file", iFileDescr); // ������� ������ ���� ���� � ������

    // ������� ��� ������ ������ ����, ���� ��� ���, �� �������
    ires = ts_file(iFileDescr, "open", filepath, "create", "we");

    if (ires != 0) {
        cout << "���� �� ������� �������:" << filepath; // ������ � ���� ���������
        return;
    }

    // ��������� � ���������� ��������� �������
    string sTableHeader = "\"ZONE_FROM_CATEGORY\";\"ZONE_FROM_NAME\";\"ZONE_TO_CATEGORY\";\"ZONE_TO_NAME\";\"DOOR_CATEGORY\"";

    ires = ts_file(iFileDescr, "write", sTableHeader+"\n"); // �������� � ���� ������

    if (ires != 0) {
        cout << "�� ������� �������� � ����";
        return;
    }


    int rowcount;
    ts_table(TableDescr1, "get_rows_count", rowcount);
    // cout << rowcount;

    for (int i = 0; i < rowcount; i++) {
        ts_table(TableDescr1, "select_row", i);
        string objectname;
        ts_table(TableDescr1, "get_value_of", 0, objectname);
        cout << objectname << "\n";

        ires = ts_file(iFileDescr, "write", objectname+"\n"); // �������� � ���� ������
        if (ires != 0) {
            cout << "�� ������� �������� � ����";
            return;
        }
    }


    ires = ts_file(iFileDescr, "close"); // ������� ����

    object("delete", iFileDescr);


    object("delete", TableDescr1); // ������� ������� � �����


    cout << "\n";
    cout << "��������� ���������� �������.\n";
    cout << "������ �� ������ ������� ������������ ���� CSV � Excel.\n";
    cout << filepath;
}

// .########.##.....##.##....##..######..########.####..#######..##....##..######.
// .##.......##.....##.###...##.##....##....##.....##..##.....##.###...##.##....##
// .##.......##.....##.####..##.##..........##.....##..##.....##.####..##.##......
// .######...##.....##.##.##.##.##..........##.....##..##.....##.##.##.##..######.
// .##.......##.....##.##..####.##..........##.....##..##.....##.##..####.......##
// .##.......##.....##.##...###.##....##....##.....##..##.....##.##...###.##....##
// .##........#######..##....##..######.....##....####..#######..##....##..######.


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