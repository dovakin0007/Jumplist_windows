use windows::{
    core::*, Win32::Foundation::*, Win32::Graphics::Gdi::ValidateRect,
    Win32::System::Com::{COINIT_APARTMENTTHREADED, COINIT_DISABLE_OLE1DDE, CoInitializeEx, CoUninitialize}, Win32::System::LibraryLoader::GetModuleHandleA,
    Win32::UI::WindowsAndMessaging::*,
};

use jumplist_win::*;

fn main() -> Result<()> {
    unsafe {
        let instance = GetModuleHandleA(None)?;
        let window_class = s!("window");
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
        let wc = WNDCLASSA {
            hCursor: LoadCursorW(None, IDC_ARROW)?,
            hInstance: instance.into(),
            lpszClassName: window_class,

            style: CS_HREDRAW | CS_VREDRAW,
            lpfnWndProc: Some(wndproc),
            ..Default::default()
        };

        let atom = RegisterClassA(&wc);
        debug_assert!(atom != 0);

        CreateWindowExA(
            WINDOW_EX_STYLE::default(),
            window_class,
            s!("This is a sample window"),
            WS_OVERLAPPEDWINDOW | WS_VISIBLE,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            None,
            None,
            instance,
            None,
        )?;

        let mut message = MSG::default();

        while GetMessageA(&mut message, None, 0, 0).into() {
            DispatchMessageA(&message);
        }
        CoUninitialize();
        Ok(())
    }
}

extern "system" fn wndproc(window: HWND, message: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    unsafe {
        match message {
            WM_CREATE => {
                println!("works");
                create_jump_list();
                LRESULT(0)
            }

            WM_PAINT => {
                println!("WM_PAINT");
                _ = ValidateRect(window, None);
                LRESULT(0)
            }
            WM_DESTROY => {
                println!("WM_DESTROY");
                PostQuitMessage(0);
                LRESULT(0)
            }
            _ => DefWindowProcA(window, message, wparam, lparam),
        }
    }
}

unsafe fn create_jump_list() {
    let mut jump_list = JumpList::new();

    // Define a common set of arguments for demonstration
    let common_args = vec![String::from("arg1"), String::from("arg2")];

    // Creating a custom category with arguments
    let mut custom_category = JumpListCategoryCustom::new(JumpListCategoryType::Custom, Some(String::from("Custom Category")));
    let mut custom_link = JumpListItemLink::new(
        Some(common_args.clone()),
        String::from("Custom Item"),
        Some(String::from("C:\\Path\\To\\Executable.exe")),
        Some(String::from("C:\\Path\\To\\Icon.ico")),
        0,
    );
    custom_link.set_working_dir(String::from("C:\\"));
    custom_category.jump_list_category.set_visible(true);
    custom_category.jump_list_category.items.push(Box::new(custom_link));
    jump_list.add_category(custom_category);

    // Creating a recent category with arguments
    let mut recent_category = JumpListCategoryCustom::new(JumpListCategoryType::Recent, None);
    let mut recent_link = JumpListItemLink::new(
        Some(common_args.clone()),
        String::from("Recent Item"),
        Some(String::from("C:\\Path\\To\\RecentExecutable.exe")),
        None,
        0,
    );
    recent_category.jump_list_category.set_visible(true);
    recent_category.jump_list_category.items.push(Box::new(recent_link));
    jump_list.add_category(recent_category);

    // Creating a frequent category with arguments
    let mut frequent_category = JumpListCategoryCustom::new(JumpListCategoryType::Frequent, None);
    let mut frequent_link = JumpListItemLink::new(
        Some(common_args.clone()),
        String::from("Frequent Item"),
        Some(String::from("C:\\Path\\To\\FrequentExecutable.exe")),
        None,
        0,
    );
    frequent_category.jump_list_category.set_visible(true);
    frequent_category.jump_list_category.items.push(Box::new(frequent_link));
    jump_list.add_category(frequent_category);

    // Creating a task category with arguments
    let mut task_category = JumpListCategoryCustom::new(JumpListCategoryType::Task, Some(String::from("Tasks")));
    let mut task_link = JumpListItemLink::new(
        Some(common_args.clone()),
        String::from("Task Item"),
        Some(String::from("C:\\Path\\To\\TaskExecutable.exe")),
        None,
        0,
    );
    task_link.set_working_dir(String::from("D:\\"));
    task_category.jump_list_category.set_visible(true);
    task_category.jump_list_category.items.push(Box::new(task_link));
    jump_list.add_category(task_category);

    jump_list.update(None);
}
