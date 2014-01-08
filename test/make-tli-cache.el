(require 'vbasense)
(require 'el-expectations)
(require 'tenv)

(expectations
  (desc "make-tli-cache not exist file")
  (expect nil
    (let* ((tdir (tenv-get-tmp-directory "vbasense" t t))
           (vbasense-cache-directory tdir)
           (vbasense-tli-files '("~/.emacs.hoge")))
      (vbasense--make-tli-cache)
      vbasense-tli-files))
  (desc "make-tli-cache not tli file")
  (expect t
    (let* ((tdir (tenv-get-tmp-directory "vbasense"))
           (vbasense-cache-directory tdir)
           (vbasense-tli-files '("~/.emacs")))
      (yaxception:$
        (yaxception:try
          (vbasense--make-tli-cache)
          nil)
        (yaxception:catch 'vbasense--ipc-err e
          (and (> (string-match "the file doesn't contain a valid typelib"
                                (yaxception:get-text e))
                  0)
               (not (file-exists-p (concat tdir "/.emacs"))))))))
  (desc "make-tli-cache tli file")
  (expect t
    (let* ((tdir (tenv-get-tmp-directory "vbasense"))
           (vbasense-cache-directory tdir)
           (vbasense-tli-files '("c:/Program Files/Common Files/Microsoft Shared/VBA/VBA6/VBE6.DLL")))
      (vbasense--make-tli-cache)
      (file-exists-p (concat tdir "/VBE6"))))
  (desc "make-tli-cache multi files")
  (expect t
    (let* ((tdir (tenv-get-tmp-directory "vbasense" t t))
           (vbasense-cache-directory tdir)
           (vbasense-tli-files '("~/.emacs.hoge"
                                 "c:/Program Files/Common Files/Microsoft Shared/VBA/VBA6/VBE6.DLL")))
      (vbasense--make-tli-cache)
      (file-exists-p (concat tdir "/VBE6"))))
  )

