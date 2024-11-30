#[macro_export]
macro_rules! get_all_queries {
    ($file:expr) => {{
        use std::collections::HashMap;
        // Helper function to load queries
        fn load_queries(sql_content: &str) -> HashMap<String, String> {
            let mut queries = HashMap::new();
            let mut current_name = String::new();
            let mut current_query = String::new();
            for line in sql_content.lines() {
                if let Some(query_line) = line.strip_prefix("-- query : ") {
                    if let Some((name, _)) = query_line.split_once('#') {
                        let trimmed_name = name.trim();
                        if !current_name.is_empty() {
                            queries.insert(current_name.clone(), current_query.trim().to_string());
                        }
                        current_name = trimmed_name.to_string();
                        current_query.clear();
                    }
                } else {
                    current_query.push_str(line);
                    current_query.push('\n');
                }
            }
            if !current_name.is_empty() {
                queries.insert(current_name, current_query.trim().to_string());
            }
            queries
        }
        let sql_content = include_str!($file);
        load_queries(sql_content)
    }};
}
