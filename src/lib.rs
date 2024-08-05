#![allow(warnings)]
use std::error::Error;
use windows::core::{GUID, Interface, PROPVARIANT, PWSTR};
use windows::Win32::System::Com::{CLSCTX_INPROC_SERVER, CoCreateInstance};
use windows::Win32::UI::Shell::{EnumerableObjectCollection, Common::IObjectCollection, IShellLinkW, ShellLink, ICustomDestinationList, DestinationList, KDC_RECENT, KDC_FREQUENT, IApplicationDestinations, SHAddToRecentDocs, SHARD_SHELLITEM, IShellItem, SHCreateItemInKnownFolder, FOLDERID_Documents, KF_FLAG_DEFAULT, ApplicationDestinations};
use windows::Win32::UI::Shell::PropertiesSystem::{IPropertyStore, PROPERTYKEY};
use windows::Win32::Storage::EnhancedStorage::PKEY_Title;
use windows::Win32::UI::Shell::Common::IObjectArray;

///TODO Add a way to create and add KnownJumplist and update it Recent and Frequent to create automatic jumplist
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
            str.push(u16::try_from('\0').unwrap());
            let new_str = PWSTR(str.as_mut_ptr());
            unsafe { link.SetPath(new_str)?; }
        }
        if let Some(args) = &self.command_args {
            let command_line = args.iter().map(|arg| format!(" \"{}\"", arg.replace('"', "\\\""))).collect::<String>();
            let mut cmd_line = String::from(command_line).encode_utf16().collect::<Vec<u16>>();
            cmd_line.push(u16::try_from('\0').unwrap());
            let command_pwstr = PWSTR(cmd_line.as_mut_ptr());
            unsafe { link.SetArguments(command_pwstr)?; }
        }
        if let Some(wd) = &self.working_dir {
            let mut wd_str = String::from(wd).encode_utf16().collect::<Vec<u16>>();
            wd_str.push(u16::try_from('\0').unwrap());
            let wd_pwstr = PWSTR(wd_str.as_mut_ptr());
            unsafe { link.SetWorkingDirectory(wd_pwstr)?; }
        }
        if let Some(icon) = &self.icon {
            let mut icon_str = String::from(icon).encode_utf16().collect::<Vec<u16>>();
            icon_str.push(u16::try_from('\0').unwrap());
            let icon_pwstr = PWSTR(icon_str.as_mut_ptr());
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
    items: Vec<Box<dyn JumpListItemTrait>>,
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
    pub fn add_item(&mut self, item: Box<dyn JumpListItemTrait>) {
        self.items.retain(|x| unsafe {
            let val = x.get_link().unwrap().cast::<IPropertyStore>().unwrap();
            let item_to_be_add = item.get_link().unwrap().cast::<IPropertyStore>().unwrap();
            if val.GetValue(&PKEY_Title).unwrap().to_string() == item_to_be_add.GetValue(&PKEY_Title).unwrap().to_string() {
                false
            }else {
                true
            }
        }
        );
        self.items.push(item);
    }
    pub unsafe fn get_category(&mut self) -> Result<IObjectCollection, Box<dyn Error>> {
        let obj_collection: *const GUID = &EnumerableObjectCollection;
        let collection: IObjectCollection = CoCreateInstance(obj_collection, None, CLSCTX_INPROC_SERVER)?;
        let mut items_to_remove = vec![];

        for (index, item) in self.items.iter().enumerate() {
            if let Ok(link) = item.get_link() {
                match self.jl_category_type {
                     JumpListCategoryType::Recent | JumpListCategoryType::Frequent =>{
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
            _ => None,
        };
        let rem_obj = match self.jumplist.GetRemovedDestinations::<IObjectArray>() {
            Ok(removed_obj) => Some(removed_obj),
            _ => None,
        };
        if let Some(rem_obj) = rem_obj {
            let rem_obj_count = rem_obj.GetCount().unwrap();
            println!("{:?}", rem_obj_count);
            for i in 0..rem_obj_count {
                if let Ok(removed_item) = rem_obj.GetAt::<IShellLinkW>(i) {
                    let removed_title = {
                        let removed_properties = removed_item.cast::<IPropertyStore>().unwrap();
                        let mut removed_propvariant = removed_properties.GetValue(&PKEY_Title).unwrap();
                        removed_propvariant.to_string()
                    };
                    self.custom.iter_mut().for_each(|category| {
                        category.jump_list_category.items.retain(|item| {
                            if let Ok(link) = item.get_link() {
                                let properties = link.cast::<IPropertyStore>().unwrap();
                                let property_store: IPropertyStore = properties;
                                let mut propvariant= property_store.GetValue(&PKEY_Title).unwrap();
                                let title = propvariant.to_string();
                                let x = title == removed_title;
                                return !x
                            } else {
                                true
                            }
                        });
                    });
                }
            }
        }


        // Remaining logic for adding the categories to the jumplist
        for category in &mut self.custom {
            match category.jump_list_category.jl_category_type {
                JumpListCategoryType::Custom => {
                    if category.jump_list_category.visible && !category.jump_list_category.items.is_empty() {
                        if let Some(title) = &category.title {
                            let mut title_wstr = String::from(title).encode_utf16().collect::<Vec<u16>>();
                            title_wstr.push(0);
                            if let Ok(value) = category.jump_list_category.get_category() {
                                if let Err(err) = self.jumplist.AppendCategory(PWSTR(title_wstr.as_ptr() as *mut u16), &value) {
                                    eprintln!("Error appending custom category: {:?}", err);
                                }
                            }
                        }
                    }
                }
                JumpListCategoryType::Recent => {
                    if let Err(err) = self.jumplist.AppendKnownCategory(KDC_RECENT) {
                        eprintln!("Error appending recent category: {:?}", err);
                    }
                }
                JumpListCategoryType::Frequent => {
                    if category.jump_list_category.visible && !category.jump_list_category.items.is_empty() {
                        if let Err(err) = self.jumplist.AppendKnownCategory(KDC_FREQUENT) {
                            eprintln!("Error appending frequent category: {:?}", err);
                        }
                    }
                }
                JumpListCategoryType::Task => {
                    if category.jump_list_category.visible && !category.jump_list_category.items.is_empty() {
                        if let Ok(value) = category.jump_list_category.get_category() {
                            if let Err(err) = self.jumplist.AddUserTasks(&value) {
                                eprintln!("Error adding user tasks: {:?}", err);
                            }
                        }
                    }
                }
            }
        }

        if let Err(err) = self.jumplist.CommitList() {
            eprintln!("Error committing jump list: {:?}", err);
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
pub unsafe fn clear_jumplist_history() {
    let dest:ICustomDestinationList = CoCreateInstance(&DestinationList, None, CLSCTX_INPROC_SERVER).unwrap();
    dest.DeleteList(None).unwrap();
}


#[cfg(test)]
mod tests {
    use super::*;
    use windows::Win32::System::Com::{COINIT_APARTMENTTHREADED, CoInitializeEx};

    #[test]
    fn test_add_category_and_item() {
        unsafe {
            CoInitializeEx(None, COINIT_APARTMENTTHREADED).unwrap();
            let mut jumplist = JumpList::new();

            let mut custom_category = JumpListCategoryCustom::new(
                JumpListCategoryType::Custom,
                Some("Custom Category".to_string())
            );

            let item_link = JumpListItemLink::new(
                Some(vec!["arg1".to_string(), "arg2".to_string()]),
                "Item Title".to_string(),
                Some("C:\\Path\\To\\Executable.exe".to_string()),
                Some("C:\\Path\\To\\Icon.ico".to_string()),
                0,
            );
            custom_category.jump_list_category.items.push(Box::new(item_link));
            jumplist.add_category(custom_category);
            assert_eq!(jumplist.custom.len(), 1, "Failed to add custom category");
            assert_eq!(jumplist.custom[0].jump_list_category.items.len(), 1, "Failed to add item to category");
        }
    }

    #[test]
    fn test_update_jumplist() {
        unsafe {
            CoInitializeEx(None, COINIT_APARTMENTTHREADED).unwrap();
            let mut jumplist = JumpList::new();

            let mut custom_category = JumpListCategoryCustom::new(
                JumpListCategoryType::Custom,
                Some("Custom Category".to_string())
            );

            let item_link = JumpListItemLink::new(
                Some(vec!["arg1".to_string(), "arg2".to_string()]),
                "Item Title".to_string(),
                Some("C:\\Path\\To\\Executable.exe".to_string()),
                Some("C:\\Path\\To\\Icon.ico".to_string()),
                0,
            );

            custom_category.jump_list_category.items.push(Box::new(item_link));
            jumplist.add_category(custom_category);

            jumplist.update();
        }
    }

    #[test]
    fn test_clear_jumplist_history() {
        unsafe {
            CoInitializeEx(None, COINIT_APARTMENTTHREADED).unwrap();
            clear_jumplist_history();
        }
    }
}



