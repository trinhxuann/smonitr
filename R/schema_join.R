#' Applies the joins in accordance to a relationship schema.
#'
#' @param schema A schema table that mirrors the structure of an Access schema.
#' @param data A list of data tables to be joined
#' @param start The starting relational table to begin construction. If NULL, will start at the first table specified by the schema.
#' @param returnSchema Return the schema table as modified by the function prior to the joining process. Internal use only.
#'
#' @return
#' @export
#'
#' @examples
schema_join <- function(schema, data, start = NULL, returnSchema = F) {

  # Account for schemas with multiple joining keys
  schema <- schema %>%
    filter(!grepl("^MSys.*", szRelationship)) %>%
    group_by(grbit, szObject, szReferencedObject,
             szRelationship = factor(szRelationship, levels = unique(szRelationship))) %>%
    summarise(foreignKeys = paste(szColumn, collapse = "."),
              primaryKeys = paste(szReferencedColumn, collapse = "."), .groups = "drop") %>%
    arrange(szRelationship)

  if (any(tolower(names(schema)) %in% "grbit")) {
    schema$joinType <- sapply(schema$grbit, function(x) {
      if (x < 16777216) "inner_join"
      else (if (x >= 16777216 & x < 33554432) "left_join"
            else ("right_join"))
    })

    schema$joinFunction <- sapply(schema$grbit, function(x) {
      if (x < 16777216) dplyr::inner_join
      else (if (x >= 16777216 & x < 33554432) dplyr::left_join
            else (dplyr::right_join))
    })
  }

  if (isTRUE(returnSchema)) return(schema)

  startTable <- data[[schema$szReferencedObject[1]]]
  usedTable <- schema$szReferencedObject[1]

  for (i in 1:nrow(schema)) {
    # If matches table already used, use the non-parent table
    if (any(schema$szObject[[i]] %in% usedTable)) {
      xTable <- schema$szObject[i]
      yTable <- schema$szReferencedObject[i]
      xName <- unlist(strsplit(schema$foreignKeys[[i]], "\\."))
      yName <- unlist(strsplit(schema$primaryKeys[[i]], "\\."))
    } else {
      xTable <- schema$szReferencedObject[i]
      yTable <- schema$szObject[i]
      xName <- unlist(strsplit(schema$primaryKeys[[i]], "\\."))
      yName <- unlist(strsplit(schema$foreignKeys[[i]], "\\."))
    }
    cat(schema$joinType[[i]], sQuote(xTable), "with", sQuote(yTable), "via columns", sQuote(xName), "and", sQuote(yName), "\n")

    startTable <- schema$joinFunction[[i]](startTable,
                                       data[[yTable]],
                                       by = setNames(yName, xName))
    usedTable <- c(usedTable, yTable)
  }
  startTable
}
