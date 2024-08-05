use windows::{
    core::*, Win32::Foundation::*, Win32::Graphics::Gdi::ValidateRect,
    Win32::System::Com::{COINIT_APARTMENTTHREADED, COINIT_DISABLE_OLE1DDE, CoInitializeEx, CoUninitialize}, Win32::System::LibraryLoader::GetModuleHandleA,
    Win32::UI::WindowsAndMessaging::*,
};
use std::ffi::{c_void, OsStr};
use std::{io, io::Result};
use std::mem::{MaybeUninit};
use std::os::windows::ffi::OsStrExt;
use std::time::Duration;

use windows::{
    Win32::Foundation::{MAX_PATH, ERROR_SUCCESS},
    Win32::System::Registry::{
        RegCloseKey, RegCreateKeyExW, RegDeleteKeyValueW, RegDeleteTreeW, RegOpenKeyW, RegSetValueExW,
        HKEY, HKEY_CLASSES_ROOT, REG_NONE, REG_SZ,
    },
    Win32::UI::Shell::{SHChangeNotify, SHCNE_ASSOCCHANGED, SHCNF_IDLIST, SHSetValueW},
    Win32::System::LibraryLoader::GetModuleFileNameW,
};
use windows::core::PCWSTR;
use windows::Win32::System::Registry::{KEY_CREATE_SUB_KEY, KEY_SET_VALUE, REG_OPTION_NON_VOLATILE};

use jumplist_win::*;

fn main() -> Result<()> {

    unsafe {
        if let Err(e) = register_to_handle_file_types() {
            eprintln!("Failed to register file types: {:?}", e);
        }
    }


    unsafe {
        if are_file_types_registered() {
            println!("File types are registered.");
        } else {
            println!("File types are not registered.");
        }
    }
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
                clear_jumplist_history();
                unsafe {
                    if let Err(e) = unregister_file_type_handlers() {
                        eprintln!("Failed to unregister file types: {:?}", e);
                    }
                }
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

macro_rules! win_err {
    ($e: expr) => {
        Err(io::Error::from_raw_os_error($e as i32))
    };
}



const PROG_ID: &str = "Microsoft.Samples.AutomaticJumpListProgID";
const EXTS_TO_REGISTER: &[&str] = &[".txt", ".doc"];

fn to_wide_str(s: &str) -> Vec<u16> {
    OsStr::new(s).encode_wide().chain(Some(0)).collect()
}

unsafe fn reg_set_string(hkey: HKEY, subkey: Option<&str>, value: Option<&str>, data: &str) -> io::Result<()> {
    let subkey_wide = subkey.map(to_wide_str);
    let value_wide = value.map(to_wide_str);
    let data_wide = to_wide_str(data);
    let len_data = ((data_wide.len() + 1) * std::mem::size_of::<u16>()) as u32;

    match SHSetValueW(
        hkey,
        subkey_wide.as_ref().map_or(None, |v| Some(PCWSTR(v.as_ptr()))).unwrap_or(PCWSTR::null()),
        value_wide.as_ref().map_or(None, |v| Some(PCWSTR(v.as_ptr()))).unwrap_or(PCWSTR::null()),
        REG_SZ.0,
        Some(data_wide.as_ptr() as *const c_void),
        len_data,
    ){
        0 => Ok(()),
        err => win_err!(err)

    }

}

unsafe fn register_progid(register: bool) -> Result<()> {
    if register {
        let mut hkey_progid = MaybeUninit::<HKEY>::zeroed();
        match RegCreateKeyExW(
            HKEY_CLASSES_ROOT,
            PCWSTR(to_wide_str(PROG_ID).as_ptr()),
            0,
            None,
            REG_OPTION_NON_VOLATILE,
            KEY_SET_VALUE | KEY_CREATE_SUB_KEY,
            None,
            hkey_progid.as_mut_ptr(),
            None,
        ).0 {
            0 => (),
            err =>return win_err!(err),
        };

        let hkey_progid = hkey_progid.assume_init();

        reg_set_string(hkey_progid, None, Some("FriendlyTypeName"), "Automatic Jump List Document")?;

        let mut app_path = vec![0u16; MAX_PATH as usize];
        let len = GetModuleFileNameW(None, &mut app_path) as usize;
        app_path.truncate(len);
        let app_path_str = String::from_utf16_lossy(&app_path);
        let icon = format!("{},0", app_path_str);

        reg_set_string(hkey_progid, Some("DefaultIcon"), None, &icon)?;
        reg_set_string(hkey_progid, Some("CurVer"), None, PROG_ID)?;

        let mut hkey_shell = MaybeUninit::<HKEY>::zeroed();
        match RegCreateKeyExW(
            hkey_progid,
            PCWSTR(to_wide_str("shell").as_ptr()),
            0,
            None,
            REG_OPTION_NON_VOLATILE,
            KEY_SET_VALUE | KEY_CREATE_SUB_KEY,
            None,
            hkey_shell.as_mut_ptr(),
            None,
        ).0 {
            0 => (),
            err => return win_err!(err)
        };

        let hkey_shell = hkey_shell.assume_init();
        let cmd_line = format!("{} %1", app_path_str);

        reg_set_string(hkey_shell, Some("Open\\Command"), None, &cmd_line)?;
        reg_set_string(hkey_shell, None, None, "Open")?;

        match RegCloseKey(hkey_shell).0 {
            0 => (),
            err => return win_err!(err)
        };
        match RegCloseKey(hkey_progid).0 {
            0 => (),
            err => return win_err!(err)
        };
    } else {
        match RegDeleteTreeW(HKEY_CLASSES_ROOT, PCWSTR(to_wide_str(PROG_ID).as_ptr())).0 {
            0 => (),
            err => return win_err!(err)
        }
    }

    Ok(())
}

unsafe fn register_to_handle_ext(ext: &str, register: bool) -> Result<()> {
    let mut key: Vec<u16> = OsStr::new(ext).encode_wide().chain(Some(0)).collect();
    key.extend(OsStr::new("\\OpenWithProgids").encode_wide().chain(Some(0)));

    let mut hkey_progid_list = MaybeUninit::<HKEY>::zeroed();
    match RegCreateKeyExW(
        HKEY_CLASSES_ROOT,
        PCWSTR(key.as_ptr()),
        0,
        None,
        REG_OPTION_NON_VOLATILE,
        KEY_SET_VALUE,
        None,
        hkey_progid_list.as_mut_ptr(),
        None,
    ).0 {
        0 => (),
        err => return win_err!(err)
    };

    let hkey_progid_list = hkey_progid_list.assume_init();

    if register {
        match RegSetValueExW(hkey_progid_list, PCWSTR(to_wide_str(PROG_ID).as_ptr()), 0, REG_NONE, None).0 {
            0 => {},
            err => return win_err!(err)
        };
    } else {
        match RegDeleteKeyValueW(hkey_progid_list, None, PCWSTR(to_wide_str(PROG_ID).as_ptr())).0 {
            0 => (),
            err => return win_err!(err)
        };

    }

    match RegCloseKey(hkey_progid_list).0 {
        0 => (),
        err => return win_err!(err)
    };
    Ok(())
}

unsafe fn register_to_handle_file_types() -> Result<()> {
    register_progid(true)?;

    for ext in EXTS_TO_REGISTER {
        register_to_handle_ext(ext, true)?;
    }

    SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, None, None);
    Ok(())
}

unsafe fn are_file_types_registered() -> bool {
    let mut hkey_progid = HKEY::default();
    let result = unsafe { RegOpenKeyW(HKEY_CLASSES_ROOT, PCWSTR(to_wide_str(PROG_ID).as_ptr()), &mut hkey_progid) };
    if result == ERROR_SUCCESS {
        RegCloseKey(hkey_progid).is_ok()
    } else {
        false
    }
}

unsafe fn unregister_file_type_handlers() -> Result<()> {
    register_progid(false)?;

    for ext in EXTS_TO_REGISTER {
        register_to_handle_ext(ext, false)?;
    }

    SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, None, None);
    Ok(())
}

// fn main() {
//     unsafe {
//         if let Err(e) = register_to_handle_file_types() {
//             eprintln!("Failed to register file types: {:?}", e);
//         }
//     }
//
//     unsafe {
//         if are_file_types_registered() {
//             println!("File types are registered.");
//         } else {
//             println!("File types are not registered.");
//         }
//     }
//
//     std::thread::sleep(Duration::from_secs(60));
//
//     unsafe {
//         if let Err(e) = unregister_file_type_handlers() {
//             eprintln!("Failed to unregister file types: {:?}", e);
//         }
//     }
// }

