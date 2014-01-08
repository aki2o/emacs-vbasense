(require 'vbasense)
(require 'el-expectations)

(expectations
  (desc "get-current-context")
  (expect ""
    (let* ((failinfo "")
           (files (directory-files (concat default-directory "sample") t "\\.\\(cls\\|bas\\|frm\\)\\'"))
           (re (rx-to-string `(and bol (* space) "'CTX:"
                                   (+ space) (group (+ (not (any space))))
                                   (+ space) (group (+ (any "0-9")))
                                   (+ space) (group (+ (any "0-9")))
                                   (+ space) (group (+ (any "a-z-")))
                                   (* space) (group (* not-newline))
                                   eol))))
      (dolist (f files)
        (vbasense-load-file f :not-update-cache t))
      (vbasense--update-availables)
      (dolist (f files)
        (vbasense--trace "Start check context : %s" f)
        (remhash (vbasense--get-tliid f) vbasense--hash-tli-file-cache)
        (with-current-buffer (find-file-noselect f)
          (vbasense-setup-current-buffer)
          (loop initially (goto-char (point-min))
                while (re-search-forward re nil t)
                for desc = (match-string-no-properties 1)
                for row = (string-to-int (match-string-no-properties 2))
                for col = (string-to-int (match-string-no-properties 3))
                for ctx-ex = (intern (match-string-no-properties 4))
                for ident-value-ex = (match-string-no-properties 5)
                do (vbasense--trace "Find context test : %s. row[%s] col[%s] ctx[%s] ident-value[%s]"
                                    desc row col ctx-ex ident-value-ex)
                do (save-excursion
                     (forward-line row)
                     (beginning-of-line)
                     (forward-char col)
                     (vbasense--update-current-definition)
                     (multiple-value-bind (ctx ident-value) (vbasense--get-current-context)
                       (when (null ident-value) (setq ident-value ""))
                       (when (or (not (eq ctx ctx-ex))
                                 (not (stringp ident-value))
                                 (not (string= ident-value ident-value-ex)))
                         (setq failinfo (concat failinfo
                                                (format "%s... Expect[%s/%s] Ret[%s/%s]\n"
                                                        desc
                                                        ctx-ex
                                                        ident-value-ex
                                                        ctx
                                                        ident-value)))))))))
      failinfo))
  )

