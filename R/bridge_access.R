#' Checks the architectures of your R and Microsoft Access programs.
#'
#' @param officeBit NULL, the architecture (32 or 64-bit) of your Microsoft Access program. If you are on Windows, this will automatically be detected; on a Linux system, you will have to provide this manually.
#'
#' @return `TRUE`/`FALSE`, where `TRUE` means that your R and Access architectures match.
#' @noRd
#' @keywords internal
architectureCheck <- function(officeBit = NULL) {

  # What architecture of R are you on?
  rBit <- ifelse((.Machine$sizeof.pointer == 4), "x32", "x64")

  # Do you have 32 bit or 64 bit office installed
  # Can attempt to read this from the registry itself;
  # if unsuccessful, the user must specify
  if (is.null(officeBit)) {
    if (rBit == "x64") {
      fp <- file.path("SOFTWARE", "Microsoft", "Office",
                      "ClickToRun", "Configuration",
                      fsep = "\\")
      subkey <- "Platform"
    } else {
      fp <- file.path("SOFTWARE", "Microsoft", "Office", "16.0", "Outlook",
                      fsep = "\\")
      subkey <- "Bitness"
    }

    officeBit <- tryCatch(readRegistry(fp)[[subkey]],
                          error = function(cond) {
                            ifelse(grepl("not found", cond$message),
                                   stop("Cannot automatically detect the architecture of your Microsoft Office. Please fill in `x32` or `x64` manually in the `officeBit` argument.", call. = F),
                                   stop(cond))
                          })
    officeBit <- ifelse((officeBit != "x64"), "x32", "x64")
  }

  # Are they the same?
  if (officeBit != rBit) {
    # First case = in 64bit R but have only 32bit office. Here, will have to use the terminal
    if (rBit == "x64" & officeBit == "x32") {
      # Check to see if a 32 bit R is installed
      if (!file.exists(paste0(Sys.getenv("R_HOME"), "/bin/i386/Rscript.exe"))) {
        stop("A 32-bit R could not be found on this machine and must be installed.", call. = F)
      }
    }
  }

  check <- ifelse(rBit == officeBit, T, F)

  list(check = check,
       rBit = rBit,
       officeBit = officeBit)
}

#' Facilitates connection from R to Access. This is meant to run on the back end.
#'
#' @param path File path to database.
#' @param driver ODBC driver. Defaults to using the Access drivers.
#' @param uid Username credential, if applicable to your database.
#' @param pwd Password credential, if applicable to your database.
#'
#' @return A DBIConnection object to allow interactions with the database.
#'
#' @noRd
#' @importFrom DBI dbConnect
#' @importFrom odbc odbc
#' @keywords internal

connectAccess <- function(path,
                           driver = "Microsoft Access Driver (*.mdb, *.accdb)", uid = "", pwd = "") {

  file <- normalizePath(path, winslash = "\\", mustWork = T)

  # Driver and path required to connect from RStudio to Access
  dbString <- paste0("Driver={", driver,
                     "};Dbq=", file,
                     ";Uid=", uid,
                     ";Pwd=", pwd,
                     ";")

  tryCatch(dbConnect(drv = odbc(), .connection_string = dbString),
           error = function(cond) {
             if (grepl(c("IM002.*ODBC Driver Manager"), cond$message)) {
               message(cond, "\n")
               message("IM002 and ODBC Driver Manager error generally means a 32-bit R needs to be installed or used.")
             } else {
               message(cond)
             }
           })
}

#' Extract tables from a connection
#'
#' @param con A DBIConnection object.
#' @param tables The tables that you wish to pull from the database. This can be left as its default, equal to "check", to return a list of tables to choose from.
#' @param out File path to store the rds file. This is required if you are on 64-bit R but have a 32-bit version of your database application, e.g., Access.
#'
#' @return A list of data tables.
#'
#' @importFrom odbc dbListTables
#' @importFrom DBI dbReadTable dbDisconnect dbGetInfo
#' @noRd
#' @keywords internal
extractTables <- function(con, tables, rBit, officeBit, out = out) {

  on.exit(dbDisconnect(con))
  # Pulling just the table names
  tableNames <- dbListTables(conn = con)

  if (length(tables) == 1 & all(tables %in% "check")) {
    # If no table names are specified, then simply return the names of the possible databases for the user to pic
    cat("Specify at least one table to pull from: \n")
    return(print(tableNames))
  }

  namesMismatch <- !tables %in% tableNames
  if (any(namesMismatch)) {
    stop(sQuote(tables[which(namesMismatch)]), " does not exist in the database.")
  }
  # Apply the dbReadTable to each readable table in db
  returnedTables <- tryCatch(mapply(dbReadTable,
                                    name = tables,
                                    MoreArgs = list(conn = con),
                                    SIMPLIFY = F),
                             error = function(cond) {
                               if (grepl("^(?=.*42000)(?=.*no read permission).*", cond$message, perl = T) &
                                   any(grepl("^MSys", tables))) {

                                 message("You are asking for a system table but do not have permission.")
                                 openDatabase <- readline(prompt = "Did you want to open the database? y/n ")

                                 if (openDatabase == "y") {
                                   shell.exec(dbGetInfo(con)$dbname)

                                   stop('Paste the command: CurrentProject.Connection.Execute "GRANT SELECT ON MSysRelationships TO Admin;". See ?bridge_access for details.', call. = F)
                                 } else {
                                   stop(cond)
                                 }
                               } else {
                                 stop(cond)
                               }
                             })

  if (rBit == "x64" & officeBit == "x32") {
    saveRDS(returnedTables, file = file.path(out, "savedAccessTables.rds"))
  } else {
    returnedTables
  }
}

#' Create the connection to an Access database and pull the requested tables. This function will attempt to account for mismatched architectures (R vs Microsoft Access) but will take longer to execute if so.
#'
#' @param file File path to the Access database file. Can be a path to a hard drive or a URL.
#' @param tables A vector of table names to determine which relational tables to pull. This can be left blank to provide a list of options. If a system table is provided, you may need to provide read permission before it will work. This has to be done in the Access DB itself, Ctrl + G, and run in the "Immediate" window: `CurrentProject.Connection.Execute "GRANT SELECT ON MSysRelationships TO Admin;"`
#'
#' @return
#' @export
#'
#' @examples
#' \dontrun{
#' bridge_access(file = "SKT.accdb")
#' bridge_access(file = "SKT.accdb", tables = c("tblCatch", "tblSample"))
#' }
#'
bridge_access <- function(file, tables = "check", ...) {

  fileType <- file(file)
  on.exit(close(fileType))

  if (class(fileType)[[1]] == "url") {
    fileName <- basename(file)

    if (!file.exists(file.path(tempdir(), fileName))) {
      download.file(file, destfile = file.path(tempdir(), fileName), mode = "wb")
    }

    if (grepl("\\.zip$", fileName)) {
      databaseName <- unzip(file.path(tempdir(), fileName), list = T)[["Name"]]
      databaseName <- databaseName[grepl("(\\.accdb)|(\\.mdb)", databaseName)]
    } else {
      databaseName <- fileName
    }

    if (length(databaseName) == 0) stop("An Access file was not found in the .zip file.")

    if (!file.exists(file.path(tempdir(), databaseName))) {
      message("Extracting file: ", sQuote(databaseName), " from the zip file.")
      unzip(file.path(tempdir(), fileName), files = databaseName, exdir = tempdir())
    }

    file <- file.path(tempdir(), databaseName)
  }

  # First, check architecture. If ok then just source the script; if not then invoke system2
  bitCheck <- architectureCheck()
  out <- tempdir()

  if (isTRUE(bitCheck$check)) {
    con <- connectAccess(file, ...)

    extractTables(con = con,
                   tables = tables,
                   rBit = bitCheck$rBit,
                   officeBit = bitCheck$officeBit,
                   out = out)
  } else {
    file <- shQuote(normalizePath(file, winslash = "\\", mustWork = T))
    script <- shQuote(normalizePath("support/connectAccessTerminal.R", winslash = "\\", mustWork = T))

    terminalOutput <- system2(paste0(Sys.getenv("R_HOME"), "/bin/i386/Rscript.exe"),
                              args = c(script,
                                       file, bitCheck, out, tables))

    # All is needed here in case length(tables) > 1 (throws warning)
    if (all(tables != "check")) readRDS(file.path(tempdir(), "savedAccessTables.rds"))
  }
}
