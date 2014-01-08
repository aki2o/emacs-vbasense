(require 'vbasense)
(require 'el-expectations)
(require 'em-glob)

(expectations
  (desc "load-recently")
  (expect t
    (let ((is-delay t)
          (not-close-file "")
          (update-count 0)
          (not-update-count 0)
          (last-update-p nil)
          (recentf-list (eshell-extended-glob (concat default-directory "sample/*"))))
      (ad-with-auto-activation-disabled
       (flet ((vbasense-load-file (file &key delay not-update-cache)
                                  (vbasense--debug "start load file. file[%s] delay[%s] not-update-cache[%s]"
                                                   file delay not-update-cache)
                                  (setq is-delay (and is-delay delay))
                                  (if not-update-cache (incf not-update-count) (incf update-count))
                                  (setq last-update-p (not not-update-cache))))
         (vbasense-load-recently)
         (vbasense--trace "pre test.\nis-delay: %s\nupdate-count: %s\nnot-update-count: %s\nlast-update-p: %s"
                          is-delay update-count not-update-count last-update-p)
         (and is-delay
              (= update-count 1)
              (> not-update-count 1)
              last-update-p)))))
  )

