use std::ptr::null_mut;
use windows::{
    core::*, Win32::Foundation::*, Win32::Graphics::Gdi::ValidateRect,
    Win32::System::Com::{COINIT_APARTMENTTHREADED, COINIT_DISABLE_OLE1DDE, CoInitializeEx, CoUninitialize}, Win32::System::LibraryLoader::GetModuleHandleA,
    Win32::UI::WindowsAndMessaging::*,
};
use windows::Win32::System::Com::CoCreateInstance;
use windows::Win32::UI::Shell::{SHAddToRecentDocs, SHARD_LINK, SHARD_SHELLITEM};
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


    // Create a Recent Items category
    let mut recent_category = JumpListCategoryCustom::new(JumpListCategoryType::Recent, None);

    // Mark the Recent category as visible
    recent_category.jump_list_category.set_visible(true);

    // Optionally, add some known items to recent
    let item_path = String::from("C:\\Users\\dovak\\OneDrive\\Documents\\message.txt");
    let shell_link = JumpListItemLink::new(
        None,
        "Recent File".to_string(),
        Some(item_path.clone()),
        None,
        0,
    );

    // Add the item to recent documents
    let shell_link_ptr = shell_link.get_link().unwrap();


    // Add recent category to JumpList
    // jump_list.add_category(recent_category);

    // Update and commit the JumpList
    // dbg!(shell_link_ptr.clone());
    jump_list.update(None);
    let mut x: [u16; 512] = [0; 512];
    &shell_link_ptr.GetPath(&mut x, null_mut(), 0);
    //println!("{:?}", x);
    let new_x = String::from_utf16_lossy(&x);
    println!("{:?}", new_x);
    SHAddToRecentDocs(SHARD_SHELLITEM.0  as u32, Some(shell_link_ptr.into_raw()));
    println!("worked  yaya");




}