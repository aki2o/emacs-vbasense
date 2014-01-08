(require 'vbasense)
(require 'el-expectations)
(require 'tenv)

(expectations
  (desc "make-progid-cache make")
  (expect t
    (let* ((tdir (tenv-get-tmp-directory "vbasense" t t))
           (vbasense-cache-directory tdir)
           (progfile (concat tdir "/" vbasense--progid-file-name)))
      (vbasense--make-progid-cache)
      (and (file-exists-p progfile)
           (> (nth 7 (file-attributes progfile)) 0))))
  (desc "make-progid-cache not make")
  (expect t
    (let* ((tdir (tenv-get-tmp-directory "vbasense"))
           (vbasense-cache-directory tdir)
           (progfile (concat tdir "/" vbasense--progid-file-name))
           (buff (find-file-noselect progfile)))
      (with-current-buffer buff
        (erase-buffer)
        (insert "a")
        (save-buffer)
        (kill-buffer))
      (vbasense--make-progid-cache)
      (and (file-exists-p progfile)
           (= (nth 7 (file-attributes progfile)) 1))))
  )

