Args <- commandArgs(T)

file <- Args[1]
rBit <- Args[3]
officeBit <- Args[4]
out <- Args[5]

tables <- Args[6:length(Args)]

con <- smonitr:::connectAccess(file)

smonitr:::extractTables(con = con,
                        tables = tables,
                        rBit = rBit,
                        officeBit = officeBit,
                        out = out)

