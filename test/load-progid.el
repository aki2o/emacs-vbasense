(require 'vbasense)
(require 'el-expectations)
(require 'tenv)

(expectations
  (desc "load-progid")
  (expect t
    (let* ((tdir (tenv-get-tmp-directory "vbasense" t t))
           (vbasense-cache-directory tdir))
      (vbasense--make-progid-cache)
      (vbasense--load-progid)
      (and (> (length vbasense--available-progids) 0)
           (= (length vbasense--available-progids)
              (length (loop for k being hash-keys in vbasense--hash-tli-prog-cache collect k))))))
  )

