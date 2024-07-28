use std::ffi::c_void;
use std::ptr::null_mut;
use windows::core::{GUID, Interface, IUnknown, PROPVARIANT, PWSTR, s};
use windows::Win32::System::Com::{CLSCTX_INPROC_SERVER, CoCreateInstance, CoInitializeEx};
use windows::Win32::UI::Shell::{EnumerableObjectCollection, Common::IObjectCollection, IShellLinkW, ShellLink, BHID_PropertyStore, ICustomDestinationList, DestinationList};
use windows::Win32::UI::Shell::PropertiesSystem::{IPropertyStore, PROPERTYKEY};
use windows::Win32::Storage::EnhancedStorage::{PKEY_AppUserModel_IsDestListSeparator, PKEY_Title};
use windows::Win32::UI::Shell::Common::IObjectArray;


pub trait JumpListItemTrait {
    fn get_link(&self) -> IShellLinkW;
}

#[derive(Default)]
pub enum JumpListItemType {
    #[default]
    Unknown = 0,
    Link = 1,
    Destination = 2,
    Separator =3,
}

pub struct JumpListItemLink {
    pub list_type: JumpListItemType,
    pub command_args: Option<Vec<String>>,
    pub title: String,
    pub command: Option<String>,
    pub icon: Option<String>,
    pub icon_index: i32,
    working_dir: Option<String>,
}


pub fn to_w_str(str: String) -> Vec<u16> {
    let mut new_str=str.encode_utf16().collect::<Vec<u16>>();
    new_str.push(0);
    new_str
}

impl JumpListItemLink {
    pub fn new(command_args: Option<Vec<String>>, title: String, command: Option<String>, icon: Option<String>, icon_index:i32) -> Self {
        Self {
            list_type: JumpListItemType::Link,
            command_args,
            title,
            command,
            icon,
            icon_index,
            working_dir: None,
        }
    }

    pub fn get_working_dir(&self) -> &Option<String>{
        &self.working_dir
    }

    pub fn set_working_dir(&mut self, wd: String) {
        self.working_dir = Some(wd);
    }
}

impl JumpListItemTrait for JumpListItemLink {
    fn get_link(&self) -> IShellLinkW {

        let shell_link_guid: *const GUID = &ShellLink;
        let link: IShellLinkW= unsafe { CoCreateInstance(shell_link_guid, None, CLSCTX_INPROC_SERVER) }.unwrap();
        let mut command =to_w_str(self.command.clone().unwrap());
        let command_pstr:PWSTR = PWSTR(command.as_mut_ptr());
        unsafe { link.SetPath(command_pstr).unwrap(); }
        let mut command_line_vec =self.command_args.clone().unwrap().iter_mut().map( |string| format!(" \"{}\"", string.replace('"', "\\\"")) ).collect::<Vec<String>>();
        let mut command_line_string = String::new();
        let _ = command_line_vec.iter_mut().map( |str|  {
            command_line_string.push_str(&str.clone()[..]);
        });
        if !command_line_vec.is_empty() {
            let command_line_string_pwstr:PWSTR = PWSTR(to_w_str( command_line_string.clone()).as_mut_ptr());
            unsafe { link.SetArguments(command_line_string_pwstr).unwrap() };
        }
        if !self.working_dir.is_none() {
            let wd_pwstr = PWSTR(to_w_str(self.working_dir.clone().unwrap()).as_mut_ptr());
            unsafe { link.SetWorkingDirectory(wd_pwstr).unwrap()};
        }
        if !self.icon.is_none() {
            let icon_pwstr = PWSTR(to_w_str(self.icon.clone().unwrap()).as_mut_ptr());
            unsafe { link.SetIconLocation(icon_pwstr, self.icon_index).unwrap() };
        }
        println!("issue");
        // let property_store_guid: *const GUID = &BHID_PropertyStore;
        let properties = link.cast::<IPropertyStore>().unwrap();

        let property_store: IPropertyStore = properties ;
        let p_title_key: *const PROPERTYKEY = &PKEY_Title ;

        let p_variant: *const PROPVARIANT = &PROPVARIANT::from(self.title.as_str());
        unsafe { property_store.SetValue(p_title_key, p_variant).unwrap(); }
        unsafe { property_store.Commit().unwrap(); }

        link

    }
}


pub struct JumpListItemSeparator {
    list_type: JumpListItemType
}

impl  JumpListItemSeparator {
    pub fn new() -> Self {
        Self {
            list_type: JumpListItemType::Separator
        }
    }

}


impl JumpListItemTrait for JumpListItemSeparator {
    fn get_link(&self) -> IShellLinkW {
        let shell_link_guid: *const GUID = &ShellLink;

        let link: IShellLinkW = unsafe { CoCreateInstance(shell_link_guid, None, CLSCTX_INPROC_SERVER) }.unwrap();
        let property_store_guid: *const GUID = &BHID_PropertyStore;
        let mut properties: *mut *mut c_void = null_mut();

        unsafe { link.query(property_store_guid, properties).unwrap() };
        let property_store: *mut IPropertyStore = properties as *mut IPropertyStore;
        let p_title_key: *const PROPERTYKEY = &PKEY_AppUserModel_IsDestListSeparator;

        unsafe {
            if let Some(property_store_back) = property_store.as_ref() {
                let p_variant: *const PROPVARIANT = &PROPVARIANT::from(true);
                property_store_back.SetValue(p_title_key, p_variant).unwrap();
                property_store_back.Commit().unwrap();
                link
            } else {
                panic!();
            }
        }
    }
}

pub enum JumpListCategoryType {
    Custom = 0,
    Task = 1,
    Recent =2,
    Frequent = 3,
}

pub struct JumpListCategory{
    pub jl_category_type: JumpListCategoryType,
    pub items: Vec<Box<dyn JumpListItemTrait>>,
    visible: bool,
}

impl JumpListCategory {
    pub fn new() -> Self {
        Self {
            jl_category_type: JumpListCategoryType::Task,
            visible: false,
            items: vec![],
        }
    }

    pub unsafe fn get_category(&self) -> IObjectCollection {
        let obj_collection:*const GUID =  &EnumerableObjectCollection;
        let clsctx_inproc = windows::Win32::System::Com::CLSCTX_INPROC_SERVER;
        let collection:IObjectCollection  = CoCreateInstance(obj_collection, None, clsctx_inproc).unwrap();
        for i in &self.items {

            collection.AddObject(&i.get_link()).unwrap();
            println!("are u even working");
        }
        collection
    }

    pub fn get_visible(&self) -> bool {
        self.visible
    }

    pub fn set_visible(&mut self, visible: bool) {
        self.visible = true
    }
}



pub struct JumpListCategoryCustom {
    pub jump_list_category: JumpListCategory,
    pub title: String
}

impl JumpListCategoryCustom {
    pub fn new(title: String) -> Self {
        let mut jl_category = JumpListCategory::new();
        jl_category.jl_category_type = JumpListCategoryType::Custom;
        Self {
            jump_list_category: jl_category,
            title
        }
    }
}

pub struct JumpList {
    jumplist: ICustomDestinationList,
    task: JumpListCategory,
    pub custom: Vec<JumpListCategoryCustom>
}

impl  JumpList {


    pub unsafe fn new() -> Self {

        let dest_list:*const GUID=&DestinationList;
        let jumplist: ICustomDestinationList = CoCreateInstance(dest_list, None, CLSCTX_INPROC_SERVER).unwrap();
        Self {
            jumplist,
            task: JumpListCategory::new(),
            custom: vec![],
        }
    }
    pub fn get_tasks(&mut self) -> &JumpListCategory {
        &self.task
    }

    pub unsafe fn update(&self) {
        //let min_slot:*mut u32 = ;

        &self.jumplist.BeginList::<IObjectArray>( &mut 10).unwrap();
        let is_command_empty = &self.task.items.is_empty().clone();


        if self.task.visible & is_command_empty {
            println!("works2");
            &self.jumplist.AddUserTasks(&self.task.get_category()).unwrap();
        }

        for category in &self.custom {
            let x = category.jump_list_category.items.is_empty();
            if category.jump_list_category.visible & !x{
                let title = to_w_str(category.title.clone());
                println!("worked");
                &self.jumplist.AppendCategory(PWSTR(title.clone().as_mut_ptr()), &category.jump_list_category.get_category()).unwrap();
            }
        }
        &self.jumplist.CommitList();
    }

    pub unsafe fn delete_list(&self, appid: String) {
        self.jumplist.SetAppID(PWSTR(to_w_str(appid.clone()).clone().as_mut_ptr())).unwrap();
        &self.jumplist.DeleteList(PWSTR(to_w_str(appid.clone()).clone().as_mut_ptr())).unwrap();
    }
    pub unsafe fn add_category(&mut self,  category: JumpListCategoryCustom) {
        &self.custom.push(category);
    }
}


