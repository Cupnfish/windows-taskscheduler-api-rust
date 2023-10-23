pub mod task_action;
pub mod task_settings;
pub mod task_trigger;

use std::mem::ManuallyDrop;

use crate::task::task_action::TaskAction;
use crate::task::task_settings::TaskSettings;
use crate::RegisteredTask;
use task_trigger::{TaskIdleTrigger, TaskLogonTrigger};

use windows::core::{ComInterface, Result};
use windows::Win32::Foundation::VARIANT_BOOL;
use windows::Win32::System::Variant::{VARIANT, VARIANT_0, VARIANT_0_0, VARIANT_0_0_0, VT_I4};

use windows::Win32::System::Com::{
    CoCreateInstance, CoInitializeEx, CLSCTX_ALL, COINIT_MULTITHREADED,
};
use windows::Win32::System::TaskScheduler::{
    IAction, IActionCollection, IExecAction, IIdleSettings, IIdleTrigger, ILogonTrigger,
    IPrincipal, IRegistrationInfo, IRepetitionPattern, ITaskDefinition, ITaskFolder,
    ITaskFolderCollection, ITaskService, ITaskSettings, ITriggerCollection, TaskScheduler,
    TASK_ACTION_EXEC, TASK_CREATE_OR_UPDATE, TASK_ENUM_HIDDEN, TASK_LOGON_INTERACTIVE_TOKEN,
    TASK_RUNLEVEL_HIGHEST, TASK_RUNLEVEL_LUA, TASK_TRIGGER_IDLE, TASK_TRIGGER_LOGON,
};

pub enum RunLevel {
    HIGHEST,
    LUA,
}

pub struct Task {
    task_definition: ITaskDefinition,
    reg_info: IRegistrationInfo,
    triggers: ITriggerCollection,
    actions: IActionCollection,
    settings: ITaskSettings,
    folder: ITaskFolder,
}
impl Task {
    fn get_task_service() -> Result<ITaskService> {
        // im probably leaking com objects memory by not releasing them but meh
        unsafe {
            CoInitializeEx(None, COINIT_MULTITHREADED)?;

            let task_service: ITaskService = CoCreateInstance(&TaskScheduler, None, CLSCTX_ALL)?;
            task_service.Connect(
                VARIANT::default(),
                VARIANT::default(),
                VARIANT::default(),
                VARIANT::default(),
            )?;
            Ok(task_service)
        }
    }

    pub fn new(path: &str) -> Result<Self> {
        unsafe {
            let task_service = Self::get_task_service()?;

            let task_definition: ITaskDefinition = task_service.NewTask(0)?;
            let triggers: ITriggerCollection = task_definition.Triggers()?;
            let reg_info: IRegistrationInfo = task_definition.RegistrationInfo()?;
            let actions: IActionCollection = task_definition.Actions()?;
            let settings: ITaskSettings = task_definition.Settings()?;
            let folder: ITaskFolder = task_service.GetFolder(&path.into())?;

            Ok(Self {
                task_definition,
                reg_info,
                triggers,
                actions,
                settings,
                folder,
            })
        }
    }

    pub fn get_registered_tasks(&self) -> Result<Vec<RegisteredTask>> {
        unsafe {
            let root_task_collection = self.folder.GetTasks(TASK_ENUM_HIDDEN.0)?;
            let mut registered_tasks = vec![];
            let count = root_task_collection.Count()?;

            for i in 0..count {
                let task = root_task_collection.get_Item(index(i))?;
                registered_tasks.push(RegisteredTask {
                    registered_task: task,
                })
            }

            let task_folder_list = self.folder.GetFolders(0)?;

            self.enum_task_folders(task_folder_list, &mut registered_tasks)?;

            Ok(registered_tasks)
        }
    }

    unsafe fn enum_task_folders(
        &self,
        task_folder_list: ITaskFolderCollection,
        registered_tasks: &mut Vec<RegisteredTask>,
    ) -> Result<()> {
        let count = task_folder_list.Count()?;
        for i in 0..count {
            let task_folder = task_folder_list.get_Item(index(i))?;
            let task_collection = task_folder.GetTasks(TASK_ENUM_HIDDEN.0)?;

            let task_collection_count = task_collection.Count()?;

            for j in 0..task_collection_count {
                let registered_task = task_collection.get_Item(index(j))?;
                registered_tasks.push(RegisteredTask { registered_task });
            }

            let tasks = self.folder.GetFolders(0)?;

            self.enum_task_folders(tasks, registered_tasks)?;
        }

        Ok(())
    }

    pub fn from_xml(self, xml: String) -> Result<Self> {
        unsafe {
            let task_service = Self::get_task_service()?;
            let task_definition: ITaskDefinition = task_service.NewTask(0)?;
            task_definition.SetXmlText(&xml.into())?;
        }
        Ok(self)
    }

    pub fn register(self, name: &str) -> Result<RegisteredTask> {
        unsafe {
            let registered_task = self.folder.RegisterTaskDefinition(
                &name.into(),
                &self.task_definition,
                TASK_CREATE_OR_UPDATE.0,
                Default::default(),
                Default::default(),
                TASK_LOGON_INTERACTIVE_TOKEN,
                Default::default(),
            )?;
            self.settings.SetEnabled(Into::<VARIANT_BOOL>::into(true))?;
            Ok(RegisteredTask { registered_task })
        }
    }

    pub fn set_hidden(self, is_hidden: bool) -> Result<Self> {
        unsafe {
            self.settings
                .SetHidden(Into::<VARIANT_BOOL>::into(is_hidden))?
        }
        Ok(self)
    }

    pub fn author(self, author: &str) -> Result<Self> {
        unsafe { self.reg_info.SetAuthor(&author.into())? }
        Ok(self)
    }

    pub fn description(self, description: &str) -> Result<Self> {
        unsafe { self.reg_info.SetDescription(&description.into())? }
        Ok(self)
    }

    pub fn idle_trigger(self, idle_trigger: TaskIdleTrigger) -> Result<Self> {
        unsafe {
            let trigger = self.triggers.Create(TASK_TRIGGER_IDLE)?;

            let i_idle_trigger: IIdleTrigger = trigger.cast::<IIdleTrigger>()?;
            i_idle_trigger.SetId(&idle_trigger.id)?;
            i_idle_trigger.SetEnabled(Into::<VARIANT_BOOL>::into(true))?;
            i_idle_trigger.SetExecutionTimeLimit(&idle_trigger.execution_time_limit)?;

            let repetition: IRepetitionPattern = i_idle_trigger.Repetition()?;
            repetition.SetInterval(&idle_trigger.repetition_interval)?;
            repetition.SetStopAtDurationEnd(Into::<VARIANT_BOOL>::into(
                idle_trigger.repetition_stop_at_duration_end,
            ))?;
        }
        Ok(self)
    }

    pub fn logon_trigger(self, logon_trigger: TaskLogonTrigger) -> Result<Self> {
        unsafe {
            let trigger = self.triggers.Create(TASK_TRIGGER_LOGON)?;
            let i_logon_trigger = trigger.cast::<ILogonTrigger>()?;
            i_logon_trigger.SetId(&logon_trigger.id)?;
            i_logon_trigger.SetEnabled(Into::<VARIANT_BOOL>::into(true))?;
            i_logon_trigger.SetExecutionTimeLimit(&logon_trigger.execution_time_limit)?;

            let repetition = i_logon_trigger.Repetition()?;
            repetition.SetInterval(&logon_trigger.repetition_interval)?;
            repetition.SetStopAtDurationEnd(Into::<VARIANT_BOOL>::into(
                logon_trigger.repetition_stop_at_duration_end,
            ))?;

            i_logon_trigger.SetDelay(&logon_trigger.delay)?;
        }
        Ok(self)
    }

    pub fn principal(self, run_level: RunLevel, id: &str, user_id: &str) -> Result<Self> {
        unsafe {
            let principal: IPrincipal = self.task_definition.Principal()?;
            match run_level {
                RunLevel::HIGHEST => principal.SetRunLevel(TASK_RUNLEVEL_HIGHEST)?,
                RunLevel::LUA => principal.SetRunLevel(TASK_RUNLEVEL_LUA)?,
            }
            principal.SetId(&id.into())?;
            principal.SetUserId(&user_id.into())?;
        }
        Ok(self)
    }

    pub fn settings(self, task_settings: TaskSettings) -> Result<Self> {
        unsafe {
            self.settings
                .SetRunOnlyIfIdle(Into::<VARIANT_BOOL>::into(task_settings.run_only_if_idle))?;
            self.settings
                .SetWakeToRun(Into::<VARIANT_BOOL>::into(task_settings.wake_to_run))?;
            self.settings
                .SetExecutionTimeLimit(&task_settings.execution_time_limit)?;
            self.settings
                .SetDisallowStartIfOnBatteries(Into::<VARIANT_BOOL>::into(
                    task_settings.disallow_start_if_on_batteries,
                ))?;
            self.settings
                .SetAllowHardTerminate(Into::<VARIANT_BOOL>::into(
                    task_settings.allow_hard_terminate,
                ))?;

            if let Some(idle_settings) = task_settings.idle_settings {
                let idle_s: IIdleSettings = self.settings.IdleSettings()?;
                idle_s
                    .SetStopOnIdleEnd(Into::<VARIANT_BOOL>::into(idle_settings.stop_on_idle_end))?;
                idle_s
                    .SetRestartOnIdle(Into::<VARIANT_BOOL>::into(idle_settings.restart_on_idle))?;
                idle_s.SetIdleDuration(&idle_settings.idle_duration)?;
                idle_s.SetWaitTimeout(&idle_settings.wait_timeout)?;
            }
        }
        Ok(self)
    }

    pub fn exec_action(self, task_action: TaskAction) -> Result<Self> {
        unsafe {
            let action: IAction = self.actions.Create(TASK_ACTION_EXEC)?;
            let exec_action: IExecAction = action.cast()?;

            exec_action.SetPath(&task_action.path)?;
            exec_action.SetId(&task_action.id)?;
            exec_action.SetWorkingDirectory(&task_action.working_dir)?;
            exec_action.SetArguments(&task_action.args)?;
        }
        Ok(self)
    }

    pub fn get_task(path: &str, name: &str) -> Result<RegisteredTask> {
        unsafe {
            let task_service = Self::get_task_service()?;
            let folder = task_service.GetFolder(&path.into())?;
            let registered_task = folder.GetTask(&name.into())?;
            Ok(RegisteredTask { registered_task })
        }
    }

    pub fn delete_task(path: &str, name: &str) -> Result<()> {
        unsafe {
            let task_service = Self::get_task_service()?;
            let folder = task_service.GetFolder(&path.into())?;
            folder.DeleteTask(&name.into(), 0)?;
        }
        Ok(())
    }
}

// struct Variant(VARIANT);
// impl Variant {
//     pub fn new(num: VARENUM, contents: VARIANT_0_0_0) -> Variant {
//         Variant {
//             0: VARIANT {
//                 Anonymous: VARIANT_0 {
//                     Anonymous: ManuallyDrop::new(VARIANT_0_0 {
//                         vt: num,
//                         wReserved1: 0,
//                         wReserved2: 0,
//                         wReserved3: 0,
//                         Anonymous: contents,
//                     }),
//                 },
//             },
//         }
//     }
// }

// impl From<String> for Variant {
//     fn from(value: String) -> Variant {
//         Variant::new(
//             VT_BSTR,
//             VARIANT_0_0_0 {
//                 bstrVal: ManuallyDrop::new(BSTR::from(value)),
//             },
//         )
//     }
// }
// impl From<&str> for Variant {
//     fn from(value: &str) -> Variant {
//         Variant::from(value.to_string())
//     }
// }
// impl From<i32> for Variant {
//     fn from(value: i32) -> Variant {
//         Variant::new(VT_I4, VARIANT_0_0_0 { lVal: value })
//     }
// }

fn index(value: i32) -> VARIANT {
    VARIANT {
        Anonymous: VARIANT_0 {
            Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                vt: VT_I4,
                wReserved1: 0,
                wReserved2: 0,
                wReserved3: 0,
                Anonymous: VARIANT_0_0_0 { lVal: value },
            }),
        },
    }
}

// impl Drop for Variant {
//     fn drop(&mut self) {
//         match VARENUM(unsafe { self.0.Anonymous.Anonymous.vt.0 }) {
//             VT_BSTR => unsafe { drop(&mut &self.0.Anonymous.Anonymous.Anonymous.bstrVal) },
//             _ => {}
//         }
//         unsafe { drop(&mut self.0.Anonymous.Anonymous) }
//     }
// }
