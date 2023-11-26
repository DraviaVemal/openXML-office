/**
 * Creates a new Word object.
 *
 * Returns a pointer to the newly created Word object.
 * If an error occurs, returns a null pointer.
 */
int8_t word_create(const char *file_name,
                   const uint8_t *buffer,
                   uintptr_t buffer_size,
                   void **out_word,
                   const char **out_error);

/**
 *Save the Word File in provided file path
 */
int8_t word_save_as(const void *word_ptr, const char *file_name, const char **out_error);

/**
 * Creates a new Power Point object.
 *
 * Returns a pointer to the newly created Power Point object.
 * If an error occurs, returns a null pointer.
 */
int8_t power_point_create(const char *file_name,
                          const uint8_t *buffer,
                          uintptr_t buffer_size,
                          void **out_power_point,
                          const char **out_error);

/**
 *Save the Power Point File in provided file path
 */
int8_t power_point_save_as(const void *power_point_ptr,
                           const char *file_name,
                           const char **out_error);

/**
 * Creates a new Excel object.
 *
 * Returns a pointer to the newly created Excel object.
 * If an error occurs, returns a null pointer.
 */
int8_t excel_create(const char *file_name,
                    const uint8_t *buffer,
                    uintptr_t buffer_size,
                    void **out_excel,
                    const char **out_error);

/**
 * Add New Sheet to the Excel
 */
int8_t excel_add_sheet(const void *excel_ptr,
                       const char *sheet_name,
                       void **out_worksheet,
                       const char **out_error);

/**
 * Get Existing Sheet from Excel
 */
int8_t excel_rename_sheet(const void *excel_ptr,
                          const char *old_sheet_name,
                          const char *new_sheet_name,
                          const char **out_error);

/**
 * Get Existing Sheet from Excel
 */
int8_t excel_get_sheet(const void *excel_ptr,
                       const char *sheet_name,
                       void **out_worksheet,
                       const char **out_error);

/**
 * List Sheet Name from Excel
 */
int8_t excel_list_sheet_name(const void *excel_ptr, const char **out_error);

/**
 *Save the Excel File in provided file path
 */
int8_t excel_save_as(const void *excel_ptr, const char *file_name, const char **out_error);
