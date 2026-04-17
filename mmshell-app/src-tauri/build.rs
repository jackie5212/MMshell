fn main() {
    #[cfg(target_os = "windows")]
    {
        cc::Build::new()
            .cpp(true)
            .file("native/drag_io.cpp")
            .file("native/virtual_drag.cpp")
            .flag_if_supported("/std:c++17")
            .compile("mmshell_drag_io");
        println!("cargo:rerun-if-changed=native/drag_io.cpp");
        println!("cargo:rerun-if-changed=native/virtual_drag.cpp");
    }
    tauri_build::build()
}
