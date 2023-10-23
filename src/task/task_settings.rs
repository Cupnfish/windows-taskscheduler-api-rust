use std::time::Duration;

use windows::core::BSTR;

pub struct IdleSettings {
    pub(crate) stop_on_idle_end: bool,
    pub(crate) restart_on_idle: bool,
    pub(crate) idle_duration: BSTR,
    pub(crate) wait_timeout: BSTR,
}
impl IdleSettings {
    pub fn new(
        stop_on_idle_end: bool,
        restart_on_idle: bool,
        idle_duration: Duration,
        wait_timeout: Duration,
    ) -> Self {
        Self {
            stop_on_idle_end,
            restart_on_idle,
            idle_duration: format!("PT{}S", idle_duration.as_secs()).into(),
            wait_timeout: format!("PT{}S", wait_timeout.as_secs()).into(),
        }
    }
}

pub struct TaskSettings {
    pub(crate) idle_settings: Option<IdleSettings>,
    pub(crate) run_only_if_idle: bool,
    pub(crate) wake_to_run: bool,
    pub(crate) execution_time_limit: BSTR,
    pub(crate) disallow_start_if_on_batteries: bool,
    pub(crate) allow_hard_terminate: bool,
}

impl TaskSettings {
    pub fn new(
        idle_settings: Option<IdleSettings>,
        run_only_if_idle: bool,
        wake_to_run: bool,
        execution_time_limit: Duration,
        disallow_start_if_on_batteries: bool,
        allow_hard_terminate: bool,
    ) -> Self {
        Self {
            idle_settings,
            run_only_if_idle,
            wake_to_run,
            execution_time_limit: format!("PT{}S", execution_time_limit.as_secs()).into(),
            disallow_start_if_on_batteries,
            allow_hard_terminate,
        }
    }
}
