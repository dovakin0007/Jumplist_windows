#![allow(warnings)]
use std::error::Error;
use windows::core::{GUID, Interface, PROPVARIANT, PWSTR};
use windows::Win32::System::Com::{CLSCTX_INPROC_SERVER, CoCreateInstance};
use windows::Win32::UI::Shell::{EnumerableObjectCollection, Common::IObjectCollection, IShellLinkW, ShellLink, ICustomDestinationList, DestinationList, KDC_RECENT, KDC_FREQUENT, IApplicationDestinations, SHAddToRecentDocs, SHARD_SHELLITEM, IShellItem, SHCreateItemInKnownFolder, FOLDERID_Documents, KF_FLAG_DEFAULT};
use windows::Win32::UI::Shell::PropertiesSystem::{IPropertyStore, PROPERTYKEY};
use windows::Win32::Storage::EnhancedStorage::PKEY_Title;
use windows::Win32::UI::Shell::Common::IObjectArray;


///TODO Add a way to create and add KnownJumplist and update it Recent and Frequent to create automatic jumplist
pub fn LOWORD(l: usize) -> usize {
    l & 0xffff
}

pub trait JumpListItemTrait {
     fn get_link(&self) -> Result<IShellLinkW, Box<dyn Error>>;
}

#[derive(Default)]
pub enum JumpListItemType {
    #[default]
    Unknown = 0,
    Link = 1,
    Destination = 2,
    Separator = 3,
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
    let mut new_str = str.encode_utf16().collect::<Vec<u16>>();
    // new_str.push(0);
    new_str
}

impl JumpListItemLink {
    pub fn new(command_args: Option<Vec<String>>, title: String, command: Option<String>, icon: Option<String>, icon_index: i32) -> Self {
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

    pub fn get_working_dir(&self) -> &Option<String> {
        &self.working_dir
    }

    pub fn set_working_dir(&mut self, wd: String) {
        self.working_dir = Some(wd);
    }
}

impl JumpListItemTrait for JumpListItemLink {
      fn get_link(&self) -> Result<IShellLinkW, Box<dyn Error>> {
        let shell_link_guid: *const GUID = &ShellLink;
        let link: IShellLinkW = unsafe { CoCreateInstance(shell_link_guid, None, CLSCTX_INPROC_SERVER)? };

        if let Some(command) = &self.command {

            let mut str = String::from(command).encode_utf16().collect::<Vec<u16>>();
            str.push(0);

            let command_pwstr = PWSTR(str.as_mut_ptr());

            unsafe { link.SetPath(command_pwstr)?; }
        }

        if let Some(args) = &self.command_args {
            let command_line = args.iter().map(|arg| format!(" \"{}\"", arg.replace('"', "\\\""))).collect::<String>();

            let command_line_pwstr = PWSTR(to_w_str(command_line).as_mut_ptr());

            unsafe { link.SetArguments(command_line_pwstr)?; }
        }

        if let Some(wd) = &self.working_dir {
            let wd_pwstr = PWSTR(to_w_str(wd.clone()).as_mut_ptr());
            unsafe { link.SetWorkingDirectory(wd_pwstr)?; }
        }

        if let Some(icon) = &self.icon {
            let icon_pwstr = PWSTR(to_w_str(icon.clone()).as_mut_ptr());
            unsafe { link.SetIconLocation(icon_pwstr, self.icon_index)?; }
        }

        let properties = link.cast::<IPropertyStore>()?;
        let property_store: IPropertyStore = properties;
        let p_title_key: *const PROPERTYKEY = &PKEY_Title;
        let p_variant: *const PROPVARIANT = &PROPVARIANT::from(self.title.as_str());
        unsafe { property_store.SetValue(p_title_key, p_variant)?; }
        unsafe { property_store.Commit()?; }

        Ok(link)
    }
}

pub struct JumpListItemSeparator {
    list_type: JumpListItemType,
}

impl JumpListItemSeparator {
    pub fn new() -> Self {
        Self {
            list_type: JumpListItemType::Separator,
        }
    }
}

impl JumpListItemTrait for JumpListItemSeparator {
     fn get_link(&self) -> Result<IShellLinkW, Box<dyn Error>> {
        let shell_link_guid: *const GUID = &ShellLink;
        let link: IShellLinkW = unsafe { CoCreateInstance(shell_link_guid, None, CLSCTX_INPROC_SERVER)? };

        let properties = link.cast::<IPropertyStore>()?;
        let property_store: IPropertyStore = properties;
        let p_title_key: *const PROPERTYKEY = &PKEY_Title;
        let p_variant: *const PROPVARIANT = &PROPVARIANT::from(true);
        unsafe { property_store.SetValue(p_title_key, p_variant)?; }
        unsafe { property_store.Commit()?; }
        Ok(link)
    }
}

pub enum JumpListCategoryType {
    Custom = 0,
    Task = 1,
    Recent = 2,
    Frequent = 3,
}

pub struct JumpListCategory {
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

    // unsafe fn check_removed_item(remove_shell_item: IObjectArray) -> bool {
    //     let no_error = &remove_shell_item.GetCount().is_ok();
    //     if *no_error {
    //         let item_count = &remove_shell_item.GetCount().unwrap();
    //         for index in item_count {
    //          match remove_shell_item.GetAt::<IShellItem>(*index){
    //             Ok(shell_item) => {
    //                shell_item.Compare()
    //             }
    //          }
    //         }
    //     }
    // }
    pub unsafe fn get_category(&mut self) -> Result<IObjectCollection, Box<dyn Error>> {
        let obj_collection: *const GUID = &EnumerableObjectCollection;
        let collection: IObjectCollection = CoCreateInstance(obj_collection, None, CLSCTX_INPROC_SERVER)?;
        let mut items_to_remove = vec![];

        for (index, item) in self.items.iter().enumerate() {
            if let Ok(link) = item.get_link() {
                match self.jl_category_type {
                     JumpListCategoryType::Recent =>{
                        SHAddToRecentDocs(SHARD_SHELLITEM.0 as u32, Some(link.as_raw()))
                    }
                    _ => {
                        collection.AddObject(&link)?;
                    }
                }

            } else {
                items_to_remove.push(index);
            }
        }

        // Remove items that failed to create links
        for index in items_to_remove.iter().rev() {
            self.items.remove(*index);
        }

        Ok(collection)
    }

    pub fn get_visible(&self) -> bool {
        self.visible
    }

    pub fn set_visible(&mut self, visible: bool) {
        self.visible = visible;
    }
}

pub struct JumpListCategoryCustom {
    pub jump_list_category: JumpListCategory,
    pub title: Option<String>,
}

impl JumpListCategoryCustom {
    pub fn new(jl_type: JumpListCategoryType, title: Option<String>) -> Self {
        match jl_type {
            JumpListCategoryType::Custom => {
                let mut jl_category = JumpListCategory::new();
                jl_category.jl_category_type = JumpListCategoryType::Custom;
                Self {
                    jump_list_category: jl_category,
                    title,
                }
            }
            JumpListCategoryType::Recent => {
                let mut jl_category = JumpListCategory::new();
                jl_category.jl_category_type = JumpListCategoryType::Recent;
                Self {
                    jump_list_category: jl_category,
                    title: None,
                }
            }
            JumpListCategoryType::Frequent => {
                let mut jl_category = JumpListCategory::new();
                jl_category.jl_category_type = JumpListCategoryType::Frequent;
                Self {
                    jump_list_category: jl_category,
                    title: None,
                }
            }
            JumpListCategoryType::Task => {
                let mut jl_category = JumpListCategory::new();
                jl_category.jl_category_type = JumpListCategoryType::Task;
                Self {
                    jump_list_category: jl_category,
                    title: None,
                }
            }
        }
    }
}

pub struct JumpList {
    jumplist: ICustomDestinationList,
    task: JumpListCategory,
    pub custom: Vec<JumpListCategoryCustom>,
}

impl JumpList {
    pub unsafe fn new() -> Self {
        let dest_list: *const GUID = &DestinationList;
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
    pub unsafe fn update(&mut self) {
        let object_array = match self.jumplist.BeginList::<IObjectArray>(&mut 10) {
            Ok(obj) => Some(obj),
            _ => None
        };
        let rem_obj = match self.jumplist.GetRemovedDestinations::<IObjectArray>() {
            Ok(removed_obj) => Some(removed_obj),
            _ => None
        };
        let rem_obj_count= &rem_obj.as_ref().unwrap().GetCount().unwrap();
        let object_array_count = &object_array.as_ref().unwrap().GetCount().unwrap();
        let mut is_present_in_removed = false;
        for i in 0..*rem_obj_count{
            for j in 0..*object_array_count{
                let obj = object_array.as_ref().unwrap();
                let rem = object_array.as_ref().unwrap();
                if obj.GetAt::<IShellLinkW>(i).unwrap() == rem.GetAt::<IShellLinkW>(j).unwrap() {
                    is_present_in_removed = true;
                }
            }
        }


        if is_present_in_removed==false {

        let mut categories_to_remove = vec![];

        for (index, category) in self.custom.iter_mut().enumerate() {
            match category.jump_list_category.jl_category_type {
                JumpListCategoryType::Custom => {
                    if category.jump_list_category.visible && !category.jump_list_category.items.is_empty() {
                        if let Some(title) = &category.title {
                            let title_wstr = to_w_str(title.clone());
                            if let Ok(mut value) = category.jump_list_category.get_category() {
                                let mut items_to_remove = vec![];
                                for (item_index, item) in category.jump_list_category.items.iter().enumerate() {
                                    if let Err(err) = item.get_link() {
                                        eprintln!("Error creating link for item: {:?}", err);
                                        items_to_remove.push(item_index);
                                    }
                                }

                                // Remove items that failed to create links
                                for index in items_to_remove.iter().rev() {
                                    category.jump_list_category.items.remove(*index);
                                }

                                if let Err(err) = self.jumplist.AppendCategory(PWSTR(title_wstr.as_ptr() as *mut u16), &value) {
                                    eprintln!("Error appending custom category: {:?}", err);
                                    categories_to_remove.push(index);
                                }
                            }
                        }
                    }
                }
                JumpListCategoryType::Recent => {
                    if category.jump_list_category.visible && !category.jump_list_category.items.is_empty() {
                        if let Err(err) = self.jumplist.AppendKnownCategory(KDC_RECENT) {
                            eprintln!("Error appending recent category: {:?}", err);
                            categories_to_remove.push(index);
                        }
                    }
                }
                JumpListCategoryType::Frequent => {
                    if category.jump_list_category.visible && !category.jump_list_category.items.is_empty() {
                        if let Err(err) = self.jumplist.AppendKnownCategory(KDC_FREQUENT) {
                            eprintln!("Error appending frequent category: {:?}", err);
                            categories_to_remove.push(index);
                        }
                    }
                }
                JumpListCategoryType::Task => {
                    if category.jump_list_category.visible && !category.jump_list_category.items.is_empty() {
                        if let Ok(mut value) = category.jump_list_category.get_category() {
                            let mut items_to_remove = vec![];
                            for (item_index, item) in category.jump_list_category.items.iter().enumerate() {
                                if let Err(err) = item.get_link() {
                                    eprintln!("Error creating link for item: {:?}", err);
                                    items_to_remove.push(item_index);
                                }
                            }

                            // Remove items that failed to create links
                            for index in items_to_remove.iter().rev() {
                                category.jump_list_category.items.remove(*index);
                            }

                            if let Err(err) = self.jumplist.AddUserTasks(&value) {
                                eprintln!("Error adding user tasks: {:?}", err);
                                categories_to_remove.push(index);
                            }
                        }
                    }
                }
            }
        }

        // Remove categories that failed to append
        for index in categories_to_remove.iter().rev() {
            self.custom.remove(*index);
        }

        if let Err(err) = self.jumplist.CommitList() {
            eprintln!("Error committing jump list: {:?}", err);
        }
        }
    }


    pub unsafe fn delete_list(&self) {
        if let Err(err) = self.jumplist.DeleteList(None) {
            eprintln!("Error deleting jump list: {:?}", err);
        }
    }

    pub unsafe fn add_category(&mut self, category: JumpListCategoryCustom) {
        self.custom.push(category);
    }
}


