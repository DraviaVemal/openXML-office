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

    pub(crate) fn get_count(
        &self,
        query: &str,
        params: &[&(dyn ToSql)],
    ) -> AnyResult<Option<usize>, AnyError> {
        fn row_mapper(row: &Row) -> Result<usize, rusqlite::Error> {
            Ok(row.get(0)?)
        }
        let mut stmt = self
            .connection
            .prepare(query)
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
    ) -> AnyResult<Option<T>, AnyError>
    where
        F: Fn(&Row) -> AnyResult<T, rusqlite::Error>,
    {
        let mut stmt = self
            .connection
            .prepare(query)
            .map_err(|e| anyhow!("Failed to Run Find One Query {}", e))?;
        match stmt.query_row(params, row_mapper) {
            Result::Ok(result) => Ok(Some(result)),
            Err(rusqlite::Error::QueryReturnedNoRows) => Ok(None),
            Err(e) => Err(e.into()),
        }
    }

    pub(crate) fn delete_record(
        &self,
        query: &str,
        params: &[&(dyn ToSql)],
    ) -> AnyResult<usize, AnyError> {
        self.connection
            .execute(query, params)
            .context("Failed to remove Record From DB")
    }

    /// Find multiple results and map each row to a specific type using the closure.
    pub(crate) fn find_many<F, T>(
        &self,
        query: &str,
        params: &[&(dyn ToSql)],
        row_mapper: F,
    ) -> AnyResult<Vec<T>, AnyError>
    where
        F: Fn(&Row) -> AnyResult<T, rusqlite::Error>,
    {
        let mut stmt = self
            .connection
            .prepare(query)
            .map_err(|e| anyhow!("Failed to Run Find Many Query {}", e))?;
        let mut results = Vec::new();
        for row in stmt.query_map(params, row_mapper)? {
            let item = row.context("Parsing the Row Error")?;
            results.push(item);
        }
        Ok(results)
    }

    /// Insert Record Into Database
    pub(crate) fn insert_record(
        &self,
        query: &str,
        params: &[&(dyn ToSql)],
    ) -> AnyResult<usize, AnyError> {
        self.connection
            .execute(&query, params)
            .context("Failed to Run Insert Record Query")
    }

    /// Insert Record Into Database
    pub(crate) fn update_record(
        &self,
        query: &str,
        params: &[&(dyn ToSql)],
    ) -> AnyResult<usize, AnyError> {
        self.connection
            .execute(&query, params)
            .context("Failed to Run Update Record Query")
    }

    /// Create Table
    pub(crate) fn create_table(&self, query: &str) -> AnyResult<usize, AnyError> {
        self.connection
            .execute(&query, ())
            .context("Failed to Run Create Table Query")
    }

    /// Insert Default Record
    pub(crate) fn insert_default(&self, query: &str) -> AnyResult<usize, AnyError> {
        self.connection
            .execute(&query, ())
            .context("Failed to Run Create Table Query")
    }

    /// Execute sent query directly without classification
    /// Try to avoid using this for code consistency
    pub(crate) fn execute_query(
        &self,
        query: &str,
        params: &[&(dyn ToSql)],
    ) -> AnyResult<usize, AnyError> {
        self.connection
            .execute(&query, params)
            .context("Failed to Run Create Table Query")
    }
}
