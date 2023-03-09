if (!exists("Args")) {
  Args <- commandArgs(T)

  file <- Args[1]
  out <- Args[2]
  tables <- Args[3:length(Args)]

  con <- connect_access(file)

  extract_tables(con = con,
                 tables = tables,
                 out = out)
}
