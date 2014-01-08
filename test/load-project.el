(require 'vbasense)
(require 'el-expectations)

(expectations
  (desc "load-project")
  (expect t
    (if (not (featurep 'anything-project))
        t
      (ap:add-project :name 'vbasense-test :look-for '(".vba-project"))
      (let ((is-delay t)
            (not-close-file "")
            (update-count 0)
            (not-update-count 0)
            (last-update-p nil))
        (ad-with-auto-activation-disabled
         (flet ((vbasense-load-file (file &key delay not-update-cache)
                                    (vbasense--debug "start load file. file[%s] delay[%s] not-update-cache[%s]"
                                                     file delay not-update-cache)
                                    (setq is-delay (and is-delay delay))
                                    (if not-update-cache (incf not-update-count) (incf update-count))
                                    (setq last-update-p (not not-update-cache))))
           (with-current-buffer (find-file-noselect (concat default-directory "sample/ISample.cls"))
             (vbasense-load-project))
           (vbasense--trace "pre test.\nis-delay: %s\nupdate-count: %s\nnot-update-count: %s\nlast-update-p: %s"
                            is-delay update-count not-update-count last-update-p)
           (and is-delay
                (= update-count 1)
                (> not-update-count 1)
                last-update-p))))))
  )

