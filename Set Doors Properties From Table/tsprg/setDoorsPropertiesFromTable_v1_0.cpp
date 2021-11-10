// ������ �������� ������ � ������������ � ����������� ������.
// LABPP AVTOMAT ������
// ���� �������
// 2021
//
//
//
// ���������� ��� � � ������� ����� � �����.xlsx�
//
// ��������� �������:
// ...
//



//--------------------------------------------------
// ������ �������� ������ �� �������
//--------------------------------------------------


string sExcelGUIDs = "���� � �����_2.xlsx";
int iTableGUIDs; // ����� ������� � GUID��� � ����������. ��������� � ������� Excel ��������� � ������ ������.
int tableRowsNumber;

int main()
{
	int ires;
	int icount;

	//--------------------------------------------------
	// ��������� ����� (���)
	//--------------------------------------------------

	// ..######..########.##.......########..######..########
	// .##....##.##.......##.......##.......##....##....##...
	// .##.......##.......##.......##.......##..........##...
	// ..######..######...##.......######...##..........##...
	// .......##.##.......##.......##.......##..........##...
	// .##....##.##.......##.......##.......##....##....##...
	// ..######..########.########.########..######.....##...


	//--------------------------------------------------
	// �������� �� ��������� ����� �� ��������� !!!!
	//--------------------------------------------------

	// ��������� ����� �� ������ ��������� ���������
	ac_request_special("load_elements_list_from_selection", 1, "DoorType", 2);

	// ��������� ��� ����� �� �������
	// ac_request("load_elements_list", 1, "DoorType", "MainFilter", 2);
	//

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

	checkAndCreateUserParameters(); // �������� ���������� � �������� �����������, ���� �����


	//--------------------------------------------------
	// ������� �������
	//--------------------------------------------------

	createTable();
	ts_table(iTableGUIDs, "get_rows_count", tableRowsNumber);
	cout << "\n";


	// ..######..##....##..######..##.......########
	// .##....##..##..##..##....##.##.......##......
	// .##.........####...##.......##.......##......
	// .##..........##....##.......##.......######..
	// .##..........##....##.......##.......##......
	// .##....##....##....##....##.##.......##......
	// ..######.....##.....######..########.########

	//--------------------------------------------------
	// ���� �� ������
	//--------------------------------------------------

	cout << "�������� �� ��������� ������. \n";
	cout << "\n";

	for (int i = 0; i < icount; i++)
	{
		ac_request("set_current_element_from_list", 1, i); // ���������� ������� ������� �� ������ � 1 � �������� i

		//--------------------------------------------------
		// ��������, � ����� ����� ��������� �������
		//--------------------------------------------------

		cout << "�������� ���� �����.\n";
		string sText, sguid, curObjFromGUID, curObjToGUID, sIDdoor, curObjFromNumber, curObjToNumber, curObjFromCat, curObjToCat, curObjFromName, curObjToName;

		ac_request("get_element_value", "GuidAsText");   // ��������� guid �������� �������� ��� �����
		sguid = ac_getstrvalue();

		// ��� ���� �������� ���� ������� �����

		curObjFromGUID = getLinkedZoneGUID (sguid, true); // �������� ���� ����: true ������ ����� �����, false � ���� ���� �����
		curObjToGUID = getLinkedZoneGUID (sguid, false);  // �������� ���� ����: true ������ ����� �����, false � ���� ���� �����

		// �������� ��� � ������ ��������� ��������� ���

		if (curObjToGUID == "") {
			cout << "   ���� _������ �� ���������.\n";
		}
		if (curObjFromGUID == "") {
			cout << "   ���� _���� �� ���������.\n";
		}
		curObjToNumber = "";
		curObjToName = "";
		curObjToCat = "";
		curObjFromNumber = "";
		curObjFromName = "";
		curObjFromCat = "";

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

		string 	curTableFromCat, curTableToCat,
		        curTableFromName, curTableToName,
		        curDoorCategory;

		bool doorWithoutCategory = true;

		// ��������� �������:
		// [ ZONE_FROM_CATEGORY | ZONE_FROM_NAME | ZONE_TO_CATEGORY | ZONE_TO_NAME | DOOR_CATEGORY ]
		// [ 7 | zone 1 | 3 | zone 5 | 015 ]

		
		ac_request("set_current_element_from_list", 1, i);
		cout << "	���������� ���������...\n";

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
			cout << "		Row: " << row+1 << " / " << tableRowsNumber << "\n";
			ts_table(iTableGUIDs, "select_row", row); // set the row

			curTableFromCat = ""; // ���������� ����������, ����� �� ������ � ��������� ��������� � ���������
			curTableFromName = "";
			curTableToCat = "";
			curTableToName = "";
			curDoorCategory = "";

			ts_table(iTableGUIDs, "get_value_of", 0, curTableFromCat);	// get current zone_from cat. in this row
			ts_table(iTableGUIDs, "get_value_of", 1, curTableFromName);	// get current zone_from name in this row
			ts_table(iTableGUIDs, "get_value_of", 2, curTableToCat);	// get current zone_from name in this row
			ts_table(iTableGUIDs, "get_value_of", 3, curTableToName);	// get current zone_from name in this row
			ts_table(iTableGUIDs, "get_value_of", 4, curDoorCategory);	// get current zone_from name in this row


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
				cout << "	���������� ��������� �����: " << ires << "\n";
				doorWithoutCategory = false;
			}

		} // end of zones for loop
			if (doorWithoutCategory == true){
				ires = ac_request("elem_user_property", "set", "DOOR_CATEGORY", "�");
				cout << "	���������� ������ ��������� �����: " << ires << "\n";
			}	
	} // end of doors for loop
	cout << "\n";
	deleteTable(); // ������� ������ �� �������
	cout << "��������� ���������� �������\n";
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

	//	runtimecontrol("workline", "setpos", 0);

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
	string sguid_linked_zone_guid = "";

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
		ts_table(iTableRoomsTmp, "get_value_of", "sGUID", sguid_linked_zone_guid);
	}
	object("delete", iGuid);
	object("delete", iTableRoomsTmp);
	return sguid_linked_zone_guid;
}

int createTable() {
	string sguid; 	// ���������� ��� guid
	string svalue; 	// ��� ���������
	int ivalue; 	// ��� ��������
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


int checkAndCreateUserParameters() {
	// ���� ���� ����� �������� �������� ���������� � ��������������, ���� �� ���

	// �� ����, ��� ���� ���������� ������� �� �������, �� �����

	// DOOR_CATEGORY � ��������� �����

	// ZONE_FROM_NAME � ��� ����, ������ ����� �����
	// ZONE_FROM_CATEGORY � ��������� ����, ������ ����� �����
	// ZONE_FROM_GUID � GUID ����, ������ ����� �����

	// ZONE_TO_NAME � ��� ����, ���� ����� �����
	// ZONE_TO_CATEGORY � ��������� ����, ���� ����� �����
	// ZONE_TO_GUID � GUID ����, ���� ����� �����

	cout << "\n��������� ����� ���������� � ������ OPENINGS.\n";

	ac_request("set_current_element_from_list", 1, 0);
	int r0, r1, r2, r3, r4, r5, r6, ires;
	bool the_res = true;

	r0 = ac_request_special("get_element_value", "UP", "DOOR_CATEGORY");
	r1 = ac_request_special("get_element_value", "UP", "ZONE_FROM_NAME");
	r2 = ac_request_special("get_element_value", "UP", "ZONE_FROM_CATEGORY");
	r3 = ac_request_special("get_element_value", "UP", "ZONE_FROM_GUID");
	r4 = ac_request_special("get_element_value", "UP", "ZONE_TO_NAME");
	r5 = ac_request_special("get_element_value", "UP", "ZONE_TO_CATEGORY");
	r6 = ac_request_special("get_element_value", "UP", "ZONE_TO_GUID");

	if (r0 == -1222) { 	//-1222 � ������ �����-��
		cout << "���� �� ������� �� �������, ���� �������.\n";
		ires = ac_request("elem_user_property", "create", "DOOR_CATEGORY", " ", "String", "OPENINGS");
		if (ires == 0) { cout << "������� �������� DOOR_CATEGORY.\n"; } else { cout << "�� ����� ������� �������� DOOR_CATEGORY, �������� �������.\n"; break();}
	}

	if (r1 == -1222) { 	//-1222 � ������ �����-��
		cout << "���� �� ������� �� �������, ���� �������.\n";
		ires = ac_request("elem_user_property", "create", "ZONE_FROM_NAME", " ", "String", "OPENINGS");
		if (ires == 0) { cout << "������� �������� ZONE_FROM_CATEGORY.\n"; } else { cout << "�� ����� ������� �������� ZONE_FROM_CATEGORY, �������� �������.\n"; break();}
	}

	if (r2 == -1222) { 	//-1222 � ������ �����-��
		cout << "���� �� ������� �� �������, ���� �������.\n";
		ires = ac_request("elem_user_property", "create", "ZONE_FROM_CATEGORY", " ", "String", "OPENINGS");
		if (ires == 0) { cout << "������� �������� ZONE_FROM_CATEGORY.\n"; } else { cout << "�� ����� ������� �������� ZONE_FROM_CATEGORY, �������� �������.\n"; break();}
	}

	if (r3 == -1222) { 	//-1222 � ������ �����-��
		cout << "���� �� ������� �� �������, ���� �������.\n";
		ires = ac_request("elem_user_property", "create", "ZONE_FROM_GUID", " ", "String", "OPENINGS");
		if (ires == 0) { cout << "������� �������� ZONE_FROM_GUID.\n"; } else { cout << "�� ����� ������� �������� ZONE_FROM_GUID, �������� �������.\n"; break();}
	}

	if (r4 == -1222) { 	//-1222 � ������ �����-��
		cout << "���� �� ������� �� �������, ���� �������.\n";
		ires = ac_request("elem_user_property", "create", "ZONE_TO_NAME", " ", "String", "OPENINGS");
		if (ires == 0) { cout << "������� �������� ZONE_TO_NAME.\n"; } else { cout << "�� ����� ������� �������� ZONE_TO_NAME, �������� �������.\n"; break();}
	}

	if (r5 == -1222) { 	//-1222 � ������ �����-��
		cout << "���� �� ������� �� �������, ���� �������.\n";
		ires = ac_request("elem_user_property", "create", "ZONE_TO_CATEGORY", " ", "String", "OPENINGS");
		if (ires == 0) { cout << "������� �������� ZONE_TO_CATEGORY.\n"; } else { cout << "�� ����� ������� �������� ZONE_TO_CATEGORY, �������� �������.\n"; break();}
	}

	if (r6 == -1222) { 	//-1222 � ������ �����-��
		cout << "���� �� ������� �� �������, ���� �������.\n";
		ires = ac_request("elem_user_property", "create", "ZONE_TO_GUID", " ", "String", "OPENINGS");
		if (ires == 0) { cout << "������� �������� ZONE_TO_GUID.\n"; } else { cout << "�� ����� ������� �������� ZONE_TO_GUID, �������� �������.\n";  break();}
	}


	cout << "\n";
	return 0;
}


