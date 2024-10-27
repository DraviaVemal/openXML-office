#[macro_export]
macro_rules! get_specific_queries {
    ($file:expr, $query_name:expr) => {{
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
        let queries = load_queries(sql_content);

        queries
            .get($query_name)
            .cloned()
            .ok_or_else(|| format!("Query '{}' not found in {}", $query_name, $file))
    }};
}

#[macro_export]
macro_rules! get_all_queries {
    ($file:expr) => {{
        let sql_content = include_str!($file);
        let mut queries = Vec::new();
        let mut current_query = String::new();

        for line in sql_content.lines() {
            if line.starts_with("-- query : ") {
                // If we hit a new query, push the previous one to the vector
                if !current_query.trim().is_empty() {
                    queries.push(current_query.trim().to_string());
                }
                current_query.clear(); // Start a new query
            } else {
                current_query.push_str(line);
                current_query.push('\n');
            }
        }

        // Push the last query if it exists
        if !current_query.trim().is_empty() {
            queries.push(current_query.trim().to_string());
        }

        queries
    }};
}
