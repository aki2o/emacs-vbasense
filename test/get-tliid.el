(require 'vbasense)
(require 'el-expectations)

(expectations
  (desc "get-tliid path")
  (expect "EXCEL"
    (vbasense--get-tliid "c:/Program Files/Microsoft Office/OFFICE11/EXCEL.EXE"))
  (desc "get-tliid version")
  (expect "vbscript_3"
    (vbasense--get-tliid "c:/WINDOWS/system32/vbscript.dll/3"))
  (desc "get-tliid tliid")
  (expect "EXCEL"
    (vbasense--get-tliid "EXCEL"))
  )

