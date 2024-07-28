// use std::env;
use untitled2::*;
// //use crate::jumplist::{JumpListItemType, JumpListItemLink, JumpListItemSeparator, JumpListCategoryType, JumpListCategory, JumpListCategoryCustom, JumpList, to_w_str};
// use std::ptr::null_mut;
// use windows::core::{Result, PCWSTR, PWSTR};
// use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, WPARAM};
// use windows::Win32::Graphics::Gdi::HBRUSH;
// use windows::Win32::System::Com::{CoInitializeEx, CoUninitialize, COINIT_APARTMENTTHREADED};
// use windows::Win32::UI::WindowsAndMessaging::{CreateWindowExW, DefWindowProcW, DispatchMessageW, GetMessageW, LoadCursorW, PostQuitMessage, RegisterClassW, TranslateMessage, CREATESTRUCTW, CW_USEDEFAULT, MSG, WM_COMMAND, WM_CREATE, WM_DESTROY, WNDCLASSW, WS_OVERLAPPEDWINDOW, WS_VISIBLE, IDC_ARROW, WINDOW_EX_STYLE};
//
// fn main() -> Result<()> {
//     unsafe {
//         env::set_var("RUST_BACKTRACE", "1");
//
//         CoInitializeEx(None, COINIT_APARTMENTTHREADED).unwrap();
//
//         let h_instance = windows::Win32::System::LibraryLoader::GetModuleHandleW(None).unwrap();
//
//
//         let class_name = PCWSTR(to_w_str(String::from("window")).as_mut_ptr());
//
//         let wnd_class = WNDCLASSW {
//             hCursor: LoadCursorW(None, IDC_ARROW).unwrap(),
//             hInstance: HINSTANCE::from(h_instance),
//             lpszClassName: class_name,
//             lpfnWndProc: Some(wnd_proc),
//             ..Default::default()
//         };
//
//
//
//         RegisterClassW(&wnd_class);
//         let lparam: *mut i32 = Box::leak(Box::new(5_i32));
//         let hwnd = CreateWindowExW(
//             WINDOW_EX_STYLE::default(),
//             class_name,
//             PCWSTR(to_w_str(String::from("JumpList example")).as_mut_ptr()),
//             WS_OVERLAPPEDWINDOW | WS_VISIBLE,
//             CW_USEDEFAULT,
//             CW_USEDEFAULT,
//             800,
//             600,
//             None,
//             None,
//             h_instance,
//             None,
//         ).unwrap();
//         println!("works");
//
//         let mut msg = MSG::default();
//         while GetMessageW(&mut msg, None, 0, 0).into() {
//             let _ = TranslateMessage(&msg);
//             DispatchMessageW(&msg);
//         }
//
//         CoUninitialize();
//     }
//
//     Ok(())
// }
//
// extern "system" fn wnd_proc(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
//     unsafe {
//         match msg {
//             WM_CREATE => {
//                 println!("works");
//                 create_jump_list();
//                 LRESULT(0)
//             }
//             WM_COMMAND => {
//                 println!("Command received: {}", wparam.0);
//                 LRESULT(0)
//             }
//             WM_DESTROY => {
//                 PostQuitMessage(0);
//                 LRESULT(0)
//             }
//             _ => DefWindowProcW(hwnd, msg, wparam, lparam),
//         }
//     }
// }
//
// unsafe fn create_jump_list() {
//     let mut jump_list = JumpList::new();
//
//     let mut task_category = JumpListCategory::new();
//     task_category.set_visible(true);
//
//     let mut item1 = JumpListItemLink::new(Some(vec!["arg1".to_string(), "arg2".to_string()]), "Task 1".to_string(), Some("notepad.exe".to_string()), None, 0);
//     item1.set_working_dir("C:\\Windows".to_string());
//     task_category.items.push(Box::new(item1));
//
//     let mut item2 = JumpListItemLink::new(Some(vec!["arg1".to_string(), "arg2".to_string()]), "Task 2".to_string(), Some("calc.exe".to_string()), None, 0);
//     task_category.items.push(Box::new(item2));
//
//     let custom_category = JumpListCategoryCustom {
//         jump_list_category: task_category,
//         title: "Tasks".to_string(),
//     };
//
//     jump_list.add_category(custom_category);
//     jump_list.update();
// }

use windows::{
    core::*, Win32::Foundation::*, Win32::Graphics::Gdi::ValidateRect,
    Win32::System::LibraryLoader::GetModuleHandleA, Win32::UI::WindowsAndMessaging::*,
};
use windows::Win32::System::Com::{COINIT_APARTMENTTHREADED, COINIT_DISABLE_OLE1DDE, CoInitializeEx, CoUninitialize};

fn main() -> Result<()> {
    unsafe {
        let instance = GetModuleHandleA(None)?;
        let window_class = s!("window");
        let _ = CoInitializeEx(None,  COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
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
            WM_CREATE =>{
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
    let mut custom_category = JumpListCategoryCustom::new(String::from("Example"));
    let args = vec![String::from("code"),String::from( ".")];
    let mut link1 = JumpListItemLink::new(
        Some(args),
        String::from("tEST_ITEM"),
        Some(String::from("works")),
        None,
        0
    );

    link1.set_working_dir(String::from("C:\\"));


    custom_category.jump_list_category.set_visible(true);
    println!("{:?}", custom_category.jump_list_category.get_visible());
    custom_category.jump_list_category.items.push(Box::new(link1));
    jump_list.add_category(custom_category);

    jump_list.update();


}

