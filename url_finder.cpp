#define UNICODE
#include <Windows.h>
#include <AtlBase.h>
#include <AtlCom.h>
#include <UIAutomation.h>

#include <iostream>
#include <locale>

/*File operations*/
#include <fstream>
#include <vector>
#include <string>

#include "resource6.h"


int ctrltab(HWND);
int ctrl_1(HWND);

void putValuesToText(std::vector<std::string> colour);


//std::locale loc("tr_TR.UTF-8");
int main()
{
    setlocale(LC_ALL, "tr_TR.UTF-8");
    CoInitialize(NULL);
    HWND hwnd = NULL;
    
    std::vector<std::string> colour;
    while (true)
    {
        hwnd = FindWindowEx(0, hwnd, L"Chrome_WidgetWin_1", NULL);
        if (!hwnd)
            break;
        if (!IsWindowVisible(hwnd))
            continue;


        ShowWindow(hwnd, 3);
        SetForegroundWindow(hwnd);


        CComQIPtr<IUIAutomation> uia;
        CComPtr<IUIAutomationElement> root;
        CComPtr<IUIAutomationCondition> chrome_condition;
        CComPtr<IUIAutomationElement> chrome_element;
        CComPtr<IUIAutomationCondition> tab_condition;
        CComPtr<IUIAutomationElement> tab_element;
        CComPtr<IUIAutomationCondition> pane_condition;
        CComPtr<IUIAutomationElement> pane_element;
        CComPtr<IUIAutomationCondition> tab_array_condition;
        CComPtr<IUIAutomationElementArray> tab_element_array;


        if (FAILED(uia.CoCreateInstance(CLSID_CUIAutomation)) || !uia)
            break;

        if (FAILED(uia->ElementFromHandle(hwnd, &root)) || !root)
            break;

        //Conditions:
        uia->CreatePropertyCondition(UIA_NamePropertyId, 
              CComVariant(L"Google Chrome"), &chrome_condition);
        uia->CreatePropertyCondition(UIA_LocalizedControlTypePropertyId,
            CComVariant(L"tab"), &tab_condition);
        uia->CreatePropertyCondition(UIA_LocalizedControlTypePropertyId,
            CComVariant(L"pane"), &pane_condition);
        uia->CreatePropertyCondition(UIA_LocalizedControlTypePropertyId,
            CComVariant(L"tab item"), &tab_array_condition);

        if (FAILED(root->FindFirst(TreeScope_Descendants, chrome_condition, &chrome_element))
            || !chrome_element)
            continue; //maybe we don't have the right chrome, continue...

        if (FAILED(chrome_element->FindFirst(TreeScope_Descendants, tab_condition, &tab_element)) || !tab_element)
            break;
        if (FAILED(tab_element->FindFirst(TreeScope_Descendants, pane_condition, &pane_element)) || !pane_element)
            break;
        if (FAILED(pane_element->FindAll(TreeScope_Children, tab_array_condition, &tab_element_array)) || !tab_element_array)
            break;

        //print now how many elements are there?
        int length;
        tab_element_array->get_Length(&length);
        CComPtr <IUIAutomationElement> elementindex;
        CComVariant urlD;
        ctrl_1(hwnd);

        

        for(int index = 0; index<length; index++){
            elementindex = NULL;
            urlD = NULL;
            tab_element_array->GetElement(index, &elementindex);
            elementindex->GetCurrentPropertyValue(UIA_NamePropertyId, &urlD);
            
            
            
            
            std::wcout << urlD.bstrVal << std::endl;

            /*NOW GET THE URL*/
            CComPtr<IUIAutomationCondition> conditionURL;
            uia->CreatePropertyCondition(UIA_NamePropertyId,
                CComVariant(L"Adres ve arama çubuðu"), &conditionURL);
            CComPtr<IUIAutomationElement> chrome_elementThis;
            if (FAILED(root->FindFirst(TreeScope_Descendants, conditionURL, &chrome_elementThis))
                || !chrome_elementThis)
                continue; //maybe we don't have the right tab, continue...

            CComVariant urlThis;
            chrome_elementThis->GetCurrentPropertyValue(UIA_ValueValuePropertyId, &urlThis);
            std::wstring ws = urlThis.bstrVal;
            std::wcout << ws << std::endl;
            std::string s(ws.begin(), ws.end());
            colour.push_back(s);
            ctrltab(hwnd);
            std::wcout << L"\n" << std::endl;
        }
        break;
    }
    CoUninitialize();
    putValuesToText(colour);
    return 0;
}
void putValuesToText(std::vector<std::string> colour) {
    std::ofstream myfile;
    myfile.open("outputs.txt");
    myfile << "Writing this to a file.\n"; //date may be added
    for (int index = 0; index < std::size(colour); index++) {
        myfile << colour[index] <<"\n"; 
        
    }
    myfile.close();
}

/// @brief This function applies "CTRL+TAB" to the chrome window 
/// @param hwnd related window handler
/// @return int 1 if no error exists
int ctrltab(HWND hwnd)
{
    if (!SetForegroundWindow(hwnd))
        return 0;


    INPUT input;
    input.type = INPUT_KEYBOARD;
    input.ki.time = 0;
    input.ki.dwExtraInfo = 0;
    input.ki.wScan = 0;
    input.ki.dwFlags = 0;

    // Ctrl Down
    input.ki.wVk = VK_CONTROL;
    SendInput(1, &input, sizeof(INPUT));
    // Tab Down 
    input.ki.wVk = VK_TAB;
    input.ki.dwFlags = 0;
    SendInput(1, &input, sizeof(INPUT));

    // Tab Up
    input.ki.dwFlags = KEYEVENTF_KEYUP;
    SendInput(1, &input, sizeof(INPUT));
    // Ctrl Up
    input.ki.wVk = VK_CONTROL;
    SendInput(1, &input, sizeof(INPUT));

    return 1;
}

int ctrl_1(HWND hwnd) {
    if (!SetForegroundWindow(hwnd))
        return 0;
    INPUT input;
    input.type = INPUT_KEYBOARD;
    input.ki.time = 0;
    input.ki.dwExtraInfo = 0;
    input.ki.wScan = 0;
    input.ki.dwFlags = 0;

    // Ctrl Down
    input.ki.wVk = VK_CONTROL;
    SendInput(1, &input, sizeof(INPUT));

    //"1" Down
    input.ki.wVk = 0x31; // virtual-key code for the "1" key
    input.ki.dwFlags = 0; // 0 for key press
    SendInput(1, &input, sizeof(INPUT));

    //"1" Up
    input.ki.dwFlags = KEYEVENTF_KEYUP; // KEYEVENTF_KEYUP for key release
    SendInput(1, &input, sizeof(INPUT));

    // Ctrl Up
    input.ki.wVk = VK_CONTROL;
    SendInput(1, &input, sizeof(INPUT));
    return 1;
}



/*References:
https://stackoverflow.com/questions/48504300/get-active-tab-url-in-chrome-with-c
//You cannot do this one:
https://stackoverflow.com/questions/65158114/chrome-automation-with-c-c-postmessage-can-send-message-even-to-a-minimized
*/
