use std::collections::HashMap;
use std::fs;

pub fn traverse_dir(dir_path: &str) -> Result<HashMap<String, String>, anyhow::Error> {
    let mut maps = HashMap::new();

    let dir = fs::read_dir(dir_path)?;
    for entry in dir {
        let entry = entry?;
        let path = entry.path();
        if path.is_file() {
            let file_name = path.file_name().unwrap().to_string_lossy().to_string();
            let file_stem = path.file_stem().unwrap().to_string_lossy().to_string();
            maps.insert(file_stem, file_name);
        } else if path.is_dir() {
            traverse_dir(&path.to_string_lossy())?;
        }
    }

    Ok(maps)
}
