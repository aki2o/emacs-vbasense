(require 'vbasense)
(require 'el-expectations)
(require 'tenv)

(expectations
  (desc "load-tli")
  (expect t
    (let* ((tdir (tenv-get-tmp-directory "vbasense" t t))
           (vbasense-cache-directory tdir)
           (vbasense-tli-files '("c:/Program Files/Common Files/Microsoft Shared/VBA/VBA6/VBE6.DLL"))
           (vbasense--hash-tli-file-cache (make-hash-table :test 'equal))
           (vbasense--hash-tli-guid-cache (make-hash-table :test 'equal))
           (vbasense--hash-tli-prog-cache (make-hash-table :test 'equal))
           (vbasense--hash-app-cache (make-hash-table :test 'equal))
           (vbasense--hash-module-cache (make-hash-table :test 'equal))
           (vbasense--hash-interface-cache (make-hash-table :test 'equal))
           (vbasense--hash-class-cache (make-hash-table :test 'equal))
           (vbasense--hash-const-cache (make-hash-table :test 'equal)))
      (vbasense--make-tli-cache)
      (vbasense--load-tli "VBE6")
      (vbasense--update-availables)
      (and (> (length vbasense--available-applications) 0)
           (> (length vbasense--available-interfaces) 0)
           (> (length vbasense--available-classes) 0)
           (> (length vbasense--available-modules) 0)
           (> (length vbasense--available-enums) 0))))
  )


