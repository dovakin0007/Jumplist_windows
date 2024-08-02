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

    // Creating a custom category for VS Code
    let mut custom_category = JumpListCategoryCustom::new(JumpListCategoryType::Custom, Some(String::from("VS Code Tasks")));

    // Directory to open in VS Code
    let directory_to_open = String::from("C:\\Users\\dovak\\OneDrive\\Documents\\lua");

    // Arguments to pass to VS Code (open the directory)
    let args = vec![directory_to_open.clone()];

    // Creating a JumpList item for VS Code
    let mut vs_code_link = JumpListItemLink::new(
        Some(args),
        String::from("Open in VS Code"),
        Some(String::from("C:\\Users\\dovak\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe")), // Path to the VS Code executable
        Some(String::from("C:\\Path\\To\\VSCodeIcon.ico")), // Optional: Path to the VS Code icon
        0,
    );

    // Set the working directory (can be the directory you're opening or any other)
    vs_code_link.set_working_dir(directory_to_open.clone());

    // Add the item to the custom category
    custom_category.jump_list_category.set_visible(true);
    custom_category.jump_list_category.items.push(Box::new(vs_code_link));

    // Add the custom category to the JumpList
    jump_list.add_category(custom_category);

    // Optionally, add other categories like recent, frequent, or tasks
    // ...

    // Update the JumpList to apply changes
    jump_list.update();
}
