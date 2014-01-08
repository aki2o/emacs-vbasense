(require 'vbasense)
(require 'el-expectations)

(expectations
  (desc "chk-tli-registration unregist")
  (expect nil
    (shell-command "regsvr32 -u c:/WINDOWS/system32/TLBINF32.DLL")
    (vbasense--chk-tli-registration))
  (desc "chk-tli-registration regist")
  (expect t
    (shell-command "regsvr32 c:/WINDOWS/system32/TLBINF32.DLL")
    (vbasense--chk-tli-registration))
  )

