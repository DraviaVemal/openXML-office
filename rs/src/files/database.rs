use anyhow::{anyhow, Context, Error as AnyError, Result as AnyResult};
use rusqlite::{Connection, Row, ToSql};
use std::{fs::remove_file, path::PathBuf};
use tempfile::env::temp_dir;

#[derive(Debug)]
pub(crate) struct SqliteDatabases {
    db_path: PathBuf,
    connection: Connection,
}

impl Drop for SqliteDatabases {
    fn drop(&mut self) {
        if self.db_path.exists() {
            match remove_file(&self.db_path) {
                Result::Ok(_) => println!(
                    "Database File successfully deleted: {}",
                    self.db_path.display()
                ),
                Err(err) => eprintln!(
                    "Failed to delete Database file {}: {}",
                    self.db_path.display(),
                    err
                ),
            }
        }
    }
}

impl SqliteDatabases {
    pub(crate) fn new(is_in_memory: bool) -> AnyResult<Self, AnyError> {
        let db_path: PathBuf = temp_dir().join(format!("{}.db", uuid::Uuid::new_v4()));
        let archive_db = Self::database_initialization(&db_path, is_in_memory)
            .context("Create Database Connection Fail")?;
        Ok(Self {
            db_path,
            connection: archive_db,
        })
    }

    /// Common initialization
    fn database_initialization(
        db_path: &PathBuf,
        is_in_memory: bool,
    ) -> AnyResult<Connection, AnyError> {
        let archive_db;
        if is_in_memory {
            // In memory operation
            archive_db = Connection::open_in_memory()
                .map_err(|e| anyhow!("Open in memory DB failed. {}", e))?;
        } else {
            archive_db = Connection::open(db_path).context("Open File Based DB failed")?;
        }
        Ok(archive_db)
    }

    /// Create Table
    pub(crate) fn create_table(
        &self,
        query: &str,
        table_name: Option<String>,
    ) -> AnyResult<usize, AnyError> {
        let mut execute_query = query.to_owned();
        if let Some(table_name) = table_name {
            execute_query = execute_query.replace("{}", &table_name);
        }
        self.connection
            .execute(&execute_query, ())
            .context("Failed to Run Create Table Query")
    }

    /// Insert Record Into Database
    pub(crate) fn insert_record(
        &self,
        query: &str,
        params: &[&(dyn ToSql)],
        table_name: Option<String>,
    ) -> AnyResult<usize, AnyError> {
        let mut execute_query = query.to_owned();
        if let Some(table_name) = table_name {
            execute_query = execute_query.replace("{}", &table_name);
        }
        self.connection
            .execute(&execute_query, params)
            .context("Failed to Run Insert Record Query")
    }

    pub(crate) fn get_count(
        &self,
        query: &str,
        params: &[&(dyn ToSql)],
        table_name: Option<String>,
    ) -> AnyResult<Option<usize>, AnyError> {
        fn row_mapper(row: &Row) -> Result<usize, rusqlite::Error> {
            Ok(row.get(0)?)
        }
        let mut execute_query = query.to_owned();
        if let Some(table_name) = table_name {
            execute_query = execute_query.replace("{}", &table_name);
        }
        let mut stmt = self
            .connection
            .prepare(&execute_query)
            .map_err(|e| anyhow!("Failed to Run Get Count Query {}", e))?;
        match stmt.query_row(params, row_mapper) {
            Result::Ok(result) => Ok(Some(result)),
            Err(rusqlite::Error::QueryReturnedNoRows) => Ok(None),
            Err(e) => Err(e.into()),
        }
    }

    /// Find one result and map the row to a specific type using the closure.
    pub(crate) fn find_one<F, T>(
        &self,
        query: &str,
        params: &[&(dyn ToSql)],
        row_mapper: F,
        table_name: Option<String>,
    ) -> AnyResult<Option<T>, AnyError>
    where
        F: Fn(&Row) -> AnyResult<T, rusqlite::Error>,
    {
        let mut execute_query = query.to_owned();
        if let Some(table_name) = table_name {
            execute_query = execute_query.replace("{}", &table_name);
        }
        let mut stmt = self
            .connection
            .prepare(&execute_query)
            .map_err(|e| anyhow!("Failed to Run Find One Query {}", e))?;
        match stmt.query_row(params, row_mapper) {
            Result::Ok(result) => Ok(Some(result)),
            Err(rusqlite::Error::QueryReturnedNoRows) => Ok(None),
            Err(e) => Err(e.into()),
        }
    }

    /// Find multiple results and map each row to a specific type using the closure.
    pub(crate) fn find_many<F, T>(
        &self,
        query: &str,
        params: &[&(dyn ToSql)],
        row_mapper: F,
        table_name: Option<String>,
    ) -> AnyResult<Vec<T>, AnyError>
    where
        F: Fn(&Row) -> AnyResult<T, rusqlite::Error>,
    {
        let mut execute_query = query.to_owned();
        if let Some(table_name) = table_name {
            execute_query = execute_query.replace("{}", &table_name);
        }
        let mut stmt = self
            .connection
            .prepare(&execute_query)
            .map_err(|e| anyhow!("Failed to Run Find Many Query {}", e))?;
        let mut results = Vec::new();
        for row in stmt.query_map(params, row_mapper)? {
            let item = row.context("Parsing the Row Error")?;
            results.push(item);
        }
        Ok(results)
    }

    pub(crate) fn delete_record(
        &self,
        query: &str,
        params: &[&(dyn ToSql)],
        table_name: Option<String>,
    ) -> AnyResult<usize, AnyError> {
        let mut execute_query = query.to_owned();
        if let Some(table_name) = table_name {
            execute_query = execute_query.replace("{}", &table_name);
        }
        self.connection
            .execute(&execute_query, params)
            .context("Failed to remove Record From DB")
    }

    pub(crate) fn drop_table(&self, table_name: String) -> AnyResult<usize, AnyError> {
        self.connection
            .execute(&format!("DROP TABLE IF EXISTS {};", table_name), ())
            .context("Failed to Drop Table")
    }

    /// Execute sent query directly without classification
    /// Try to avoid using this for code consistency
    pub(crate) fn _execute_query(
        &self,
        query: &str,
        params: &[&(dyn ToSql)],
    ) -> AnyResult<usize, AnyError> {
        self.connection
            .execute(&query, params)
            .context("Failed to Run Create Table Query")
    }
}
