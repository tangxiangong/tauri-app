use crate::utils::format_bytes;
use serde::Serialize;
use sysinfo::{IS_SUPPORTED_SYSTEM, System};

pub fn is_supported() -> bool {
    IS_SUPPORTED_SYSTEM
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Memory {
    total_memory: u64,
    used_memory: u64,
    total_swap: u64,
    used_swap: u64,
}

impl std::fmt::Display for Memory {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(
            f,
            "Total Memory: {}, Used Memory: {}, Total Swap: {}, Used Swap: {}",
            format_bytes(self.total_memory),
            format_bytes(self.used_memory),
            format_bytes(self.total_swap),
            format_bytes(self.used_swap)
        )
    }
}

pub fn get_memory() -> Memory {
    let mut sys = System::new();
    sys.refresh_all();
    Memory {
        total_memory: sys.total_memory(),
        used_memory: sys.used_memory(),
        total_swap: sys.total_swap(),
        used_swap: sys.used_swap(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_is_supported() {
        assert!(is_supported());
    }

    #[test]
    fn test_get_memory() {
        let memory = get_memory();
        println!("{memory}");
    }
}
