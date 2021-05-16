#define UNICODE
#include <Windows.h>
#include <AtlBase.h>
#include <AtlCom.h>
#include <UIAutomation.h>

#include <iostream>

/*File operations*/
#include <iostream>
#include <fstream>

int ctrltab(HWND);

int main()
{
        
    
    CoInitialize(NULL);
    HWND hwnd = NULL;
    while (true)
    {
        hwnd = FindWindowEx(0, hwnd, L"Chrome_WidgetWin_1", NULL);
        if (!hwnd)
            break;
        if (!IsWindowVisible(hwnd))
            continue;


        ShowWindow(hwnd, 3);
        SetForegroundWindow(hwnd);


        //uia ayarý
        CComQIPtr<IUIAutomation> uia;
        if (FAILED(uia.CoCreateInstance(CLSID_CUIAutomation)) || !uia)
            break;

        //root ayarý
        CComPtr<IUIAutomationElement> root;
        if (FAILED(uia->ElementFromHandle(hwnd, &root)) || !root)
            break;

        //condition ayarý: 
        CComPtr<IUIAutomationCondition> condition;
        uia->CreatePropertyCondition(UIA_NamePropertyId, 
              CComVariant(L"Google Chrome"), &condition);

        



        //edit ayarý, ikinci kez IUIAutomationElement
        CComPtr<IUIAutomationElement> edit;
        if (FAILED(root->FindFirst(TreeScope_Descendants, condition, &edit))
            || !edit)
            continue; //maybe we don't have the right tab, continue...

        //condition ayarý: 
        CComPtr<IUIAutomationCondition> condition2;
        uia->CreatePropertyCondition(UIA_LocalizedControlTypePropertyId,
            CComVariant(L"tab"), &condition2);

        CComPtr<IUIAutomationElement> second;
        if (FAILED(edit->FindFirst(TreeScope_Descendants, condition2, &second)) || !second)
            break;


        //tabdan ilk pane e geçiþ 
        CComPtr<IUIAutomationCondition> condition3;
        uia->CreatePropertyCondition(UIA_LocalizedControlTypePropertyId,
            CComVariant(L"pane"), &condition3);

        CComPtr<IUIAutomationElement> second2;
        if (FAILED(second->FindFirst(TreeScope_Descendants, condition3, &second2)) || !second2)
            break;



        CComVariant url;
        second2->GetCurrentPropertyValue(UIA_ProcessIdPropertyId, &url);
        
        //no loop
        CComPtr<IUIAutomationCondition> condition4;
        uia->CreatePropertyCondition(UIA_LocalizedControlTypePropertyId,
                CComVariant(L"tab item"), &condition4);

        CComPtr<IUIAutomationElementArray> second4;
        if (FAILED(second2->FindAll(TreeScope_Children, condition4, &second4)) || !second4)
            break;

        //print now how many elements are there?
        int length;
        second4->get_Length(&length);
        CComPtr <IUIAutomationElement> elementindex;
        CComVariant urlD;
        for(int index = 0; index<length; index++){
            elementindex = NULL;
            urlD = NULL;
            second4->GetElement(index, &elementindex);
            elementindex->GetCurrentPropertyValue(UIA_NamePropertyId, &urlD);
            std::wcout << urlD.bstrVal << std::endl;

            /*NOW GET THE URL*/
            CComPtr<IUIAutomationCondition> conditionURL;
            uia->CreatePropertyCondition(UIA_NamePropertyId,
                CComVariant(L"Adres ve arama çubuðu"), &conditionURL);
            CComPtr<IUIAutomationElement> editThis;
            if (FAILED(root->FindFirst(TreeScope_Descendants, conditionURL, &editThis))
                || !editThis)
                continue; //maybe we don't have the right tab, continue...

            CComVariant urlThis;
            editThis->GetCurrentPropertyValue(UIA_ValueValuePropertyId, &urlThis);
            std::wcout << urlThis.bstrVal << std::endl;
            ctrltab(hwnd);
            //hwnd = SendMessage(hwnd, );
            std::wcout << L"\n" << std::endl;
        }
        break;
    }
    CoUninitialize();
    return 0;
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





/*References:
https://stackoverflow.com/questions/48504300/get-active-tab-url-in-chrome-with-c
//You cannot:
https://stackoverflow.com/questions/65158114/chrome-automation-with-c-c-postmessage-can-send-message-even-to-a-minimized
*/
