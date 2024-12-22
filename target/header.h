#define ThemePalletModel_VT_ACCENT1 4

#define ThemePalletModel_VT_ACCENT2 6

#define ThemePalletModel_VT_ACCENT3 8

#define ThemePalletModel_VT_ACCENT4 10

#define ThemePalletModel_VT_ACCENT5 12

#define ThemePalletModel_VT_ACCENT6 14

#define ThemePalletModel_VT_DARK1 16

#define ThemePalletModel_VT_DARK2 18

#define ThemePalletModel_VT_LIGHT1 20

#define ThemePalletModel_VT_LIGHT2 22

#define ThemePalletModel_VT_HYPERLINK 24

#define ThemePalletModel_VT_FOLLOWED_HYPERLINK 26

#define CorePropertiesModel_VT_TITLE 4

#define CorePropertiesModel_VT_SUBJECT 6

#define CorePropertiesModel_VT_DESCRIPTION 8

#define CorePropertiesModel_VT_TAGS 10

#define CorePropertiesModel_VT_CATEGORY 12

#define CorePropertiesModel_VT_CREATOR 14

#define ExcelPropertiesModel_VT_IS_IN_MEMORY 4

#define ExcelPropertiesModel_VT_IS_EDITABLE 6

#define ExcelPropertiesModel_VT_SETTINGS 8

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
