;;; vbasense.el --- provide a environment like Visual Basic Editor.

;; Copyright (C) 2013  Hiroaki Otsu

;; Author: Hiroaki Otsu <ootsuhiroaki@gmail.com>
;; Keywords: vba, completion
;; URL: https://github.com/aki2o/emacs-vbasense
;; Version: 0.1.1
;; Package-Requires: ((auto-complete "1.4.0") (log4e "0.2.0") (yaxception "0.1"))

;; This program is free software; you can redistribute it and/or modify
;; it under the terms of the GNU General Public License as published by
;; the Free Software Foundation, either version 3 of the License, or
;; (at your option) any later version.

;; This file is distributed in the hope that it will be useful,
;; but WITHOUT ANY WARRANTY; without even the implied warranty of
;; MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
;; GNU General Public License for more details.

;; You should have received a copy of the GNU General Public License
;; along with this program.  If not, see <http://www.gnu.org/licenses/>.

;;; Commentary:
;; 
;; This extension provides a environment with the following function like Visual Basic Editor.
;; - code assist
;; - jump to definition
;; - popup help
;; - show signature of method
;; 
;; For more infomation,
;; see <https://github.com/aki2o/emacs-vbasense/blob/master/README.md>

;;; Dependency:
;; 
;; - auto-complete.el ( see <https://github.com/auto-complete/auto-complete> )
;; - yaxception.el ( see <https://github.com/aki2o/yaxception> )
;; - log4e.el ( see <https://github.com/aki2o/log4e> )

;;; Installation:
;;
;; Put this to your load-path.
;; And put the following lines in your .emacs or site-start.el file.
;; 
;; (require 'vbasense)

;;; Configuration:
;; 
;; ;; Key Binding
;; (setq vbasense-popup-help-key "C-:")
;; (setq vbasense-jump-to-definition-key "C->")
;; 
;; ;; Make config suit for you. About the config item, see Customization or eval the following sexp.
;; ;; (customize-group "vbasense")
;; ;; 
;; ;; The following items are particularly important
;; ;; ** For loading installed library, set `vbasense-tli-files'
;; ;; (add-to-list 'vbasense-tli-files "c:/Program Files/Common Files/System/ado/msado21.tlb")
;; 
;; ;; ** For loading user library, which is information of VBA source file other than the opened buffer
;; ;; - You can load from your library directory by following configuration
;; ;; (add-to-list 'vbasense-user-library-directories "c:/MyVBALibraray")
;; ;; - Otherwise, you can load more flexibly by customizing `vbasense-setup-user-library-function'
;; ;;   - If you want to load recently files, use `vbasense-load-recently'
;; ;;   - If you want to use anything-project.el, use `vbasense-load-project'
;; ;; (setq vbasense-setup-user-library-function (lambda () (vbasense-load-project)))
;; 
;; (vbasense-config-default)

;;; Customization:
;; 
;; [EVAL] (autodoc-document-lisp-buffer :type 'user-variable :prefix "vbasense-[^\-]" :docstring t)
;; `vbasense-tli-files'
;; List of TypeLibInfo file which you activate on the reference configuration of Visual Basic Editor.
;; `vbasense-cache-directory'
;; Path of directory which the infomation of TypeLibInfo file is saved in.
;; `vbasense-user-library-directories'
;; List of directory which has the user library.
;; `vbasense-setup-user-library-function'
;; Function for loading the user library when open the buffer of VBA source file.
;; `vbasense-popup-help-key'
;; Keystroke for popup help about something at point.
;; `vbasense-jump-to-definition-key'
;; Keystroke for jump to definition at point.
;; `vbasense-enable-modes'
;; Major modes vbasense is enabled on.
;; `vbasense-ac-trigger-command-keys'
;; Keystrokes for doing `ac-start' with self insert.
;; `vbasense-ac-implicit-enummember-prefixes'
;; Prefixes of the Enum member which you let be candidate of auto-complete without Enum.
;; `vbasense-lookup-current-buffer-threshold'
;; Size of buffer as threshold for looking up Procedure/Enum/Type in current buffer for each of completion.
;; 
;;  *** END auto-documentation

;;; API:
;; 
;; [EVAL] (autodoc-document-lisp-buffer :type 'command :prefix "vbasense-[^\-]" :docstring t)
;; `vbasense-load-library'
;; Load library in `vbasense-cache-directory'.
;; `vbasense-remake-tli-cache'
;; Remake library of TLIFILE.
;; `vbasense-remake-progid-cache'
;; Remake information of ProgID.
;; `vbasense-delete-all-cache'
;; Delete `vbasense-cache-directory'.
;; `vbasense-load-file'
;; Load user file.
;; `vbasense-load-directory'
;; Load user file in directory.
;; `vbasense-load-recently'
;; Load user file in recentf.
;; `vbasense-load-project'
;; Load user file in project.
;; `vbasense-reload-current-buffer'
;; Reload user file of current buffer.
;; `vbasense-popup-help'
;; Popup help about something at point.
;; `vbasense-jump-to-definition'
;; Jump to definition at point.
;; `vbasense-insert-implement-definition'
;; Insert definition of procedure which the implemented interface has.
;; `vbasense-setup-current-buffer'
;; Do setup for using vbasense in current buffer.
;; 
;;  *** END auto-documentation
;; [Note] Functions and variables other than listed above, Those specifications may be changed without notice.

;;; Tested On:
;; 
;; - Emacs ... GNU Emacs 24.2.1 (i386-mingw-nt5.1.2600) of 2012-12-08 on GNUPACK
;; - auto-complete.el ... Version 1.4.0
;; - yaxception.el ... Version 0.1
;; - log4e.el ... Version 0.2.0


;; Enjoy!!!


(eval-when-compile (require 'cl))
(require 'rx)
(require 'regexp-opt)
(require 'auto-complete)
(require 'em-glob)
(require 'recentf)
(require 'eldoc)
(require 'ring)
(require 'etags)
(require 'log4e)
(require 'yaxception)
(require 'anything-project nil t)
(require 'pos-tip nil t)


(defgroup vbasense nil
  "Auto completion for VBA source file."
  :group 'completion
  :prefix "vbasense-")

(defcustom vbasense-tli-files '("c:/Program Files/Microsoft Office/OFFICE11/EXCEL.EXE"
                                "c:/Program Files/Common Files/Microsoft Shared/VBA/VBA6/VBE6.DLL"
                                "c:/Program Files/Common Files/Microsoft Shared/VBA/VBA6/VBE6EXT.OLB"
                                "c:/Program Files/Common Files/Microsoft Shared/OFFICE11/MSO.DLL"
                                "c:/WINDOWS/system32/stdole2.tlb")
  "List of TypeLibInfo file which you activate on the reference configuration of Visual Basic Editor."
  :type '(repeat file)
  :group 'vbasense)

(defcustom vbasense-cache-directory "~/.vbasense.d"
  "Path of directory which the infomation of TypeLibInfo file is saved in."
  :type 'directory
  :group 'vbasense)

(defcustom vbasense-user-library-directories nil
  "List of directory which has the user library.

If you customize `vbasense-setup-user-library-function', this value may be no need."
  :type '(repeat directory)
  :group 'vbasense)

(defcustom vbasense-setup-user-library-function
  (lambda ()
    (dolist (dir vbasense-user-library-directories)
      (vbasense-load-directory dir)))
  "Function for loading the user library when open the buffer of VBA source file.

The user library is the information of VBA source file other than the opened buffer."
  :type 'function
  :group 'vbasense)

(defcustom vbasense-popup-help-key nil
  "Keystroke for popup help about something at point."
  :type 'string
  :group 'vbasense)

(defcustom vbasense-jump-to-definition-key nil
  "Keystroke for jump to definition at point."
  :type 'string
  :group 'vbasense)

(defcustom vbasense-enable-modes '(visual-basic-mode)
  "Major modes vbasense is enabled on."
  :type '(repeat symbol)
  :group 'vbasense)

(defcustom vbasense-ac-trigger-command-keys '("SPC" "." "," "(")
  "Keystrokes for doing `ac-start' with self insert."
  :type '(repeat string)
  :group 'vbasense)

(defcustom vbasense-ac-implicit-enummember-prefixes '("vb" "xl")
  "Prefixes of the Enum member which you let be candidate of auto-complete without Enum."
  :type '(repeat string)
  :group 'vbasense)

(defcustom vbasense-lookup-current-buffer-threshold 10000
  "Size of buffer as threshold for looking up Procedure/Enum/Type in current buffer for each of completion.

If this value is larger than 0 and current buffer size is larger than this value, do not looking up.
If this value is 0, do looking up always.

Looking up in big size buffer may cause slowness of Emacs."
  :type 'number
  :group 'vbasense)

(log4e:deflogger "vbasense" "%t [%l] %m" "%H:%M:%S" '((fatal . "fatal")
                                                      (error . "error")
                                                      (warn  . "warn")
                                                      (info  . "info")
                                                      (debug . "debug")
                                                      (trace . "trace")))
(vbasense--log-set-level 'trace)


(defvar vbasense--regexp-variable-ident "[a-zA-Z][a-zA-Z_0-9]*")
(defvar vbasense--regexp-namespace-ident "[a-zA-Z][a-zA-Z_0-9.]*")
(defvar vbasense--regexp-argument (rx-to-string `(and "("
                                                      (group (or ""
                                                                 (and (+? anything) (not (any "(")))))
                                                      ")")))

(defvar vbasense--builtin-keywords nil)
(defvar vbasense--builtin-objects '("String" "Integer" "Long" "Single" "Double" "Currency" "Date" "Any"
                                    "Object" "Boolean" "Variant" "DataObject" "Decimal" "Byte" "Array"))
(defvar vbasense--builtin-functions '("Left" "Mid" "Right" "CreateObject" "UBound" "Space" "Environ"))

(defmacro* vbasense--define-word-regexp (name &rest words)
  (declare (indent 0))
  (let* ((exwords (loop for e in words
                        append (list e (downcase e))))
         (re (rx-to-string `(and bos (group (or ,@exwords)) eos))))
    `(progn
       (loop for w in ',words
             do (pushnew w vbasense--builtin-keywords :test 'equal))
       (defvar ,(intern (concat "vbasense--words-" name)) ',exwords)
       (defvar ,(intern (concat "vbasense--regexp-word-of-" name)) ,re))))

(vbasense--define-word-regexp "scope" "Public" "Private" "Friend")
(vbasense--define-word-regexp "function" "Function")
(vbasense--define-word-regexp "sub" "Sub")
(vbasense--define-word-regexp "property" "Property")
(vbasense--define-word-regexp "get" "Get")
(vbasense--define-word-regexp "let" "Let")
(vbasense--define-word-regexp "type" "Type")
(vbasense--define-word-regexp "enum" "Enum")
(vbasense--define-word-regexp "end" "End")
(vbasense--define-word-regexp "exit" "Exit")
(vbasense--define-word-regexp "dim" "Dim")
(vbasense--define-word-regexp "static" "Static")
(vbasense--define-word-regexp "const" "Const")
(vbasense--define-word-regexp "as" "As")
(vbasense--define-word-regexp "new" "New")
(vbasense--define-word-regexp "set" "Set")
(vbasense--define-word-regexp "createobject" "CreateObject")
(vbasense--define-word-regexp "for" "For")
(vbasense--define-word-regexp "each" "Each")
(vbasense--define-word-regexp "in" "In")
(vbasense--define-word-regexp "class" "As" "New")
(vbasense--define-word-regexp "implement" "Implements")
(vbasense--define-word-regexp "variable" "Set" "ReDim" "Preserve")
(vbasense--define-word-regexp "optional" "Optional")
(vbasense--define-word-regexp "procedure-modifier" "Static" "Declare")
(vbasense--define-word-regexp "definition" "Function" "Sub" "Get" "Let" "Type" "Enum" "Dim" "Const" "ByVal" "ByRef" "ParamArray" "Lib" "Alias")
(vbasense--define-word-regexp "label" "GoTo")

(loop for w in '("If" "Then" "ElseIf" "Else" "For" "Next" "To" "Step" "Each" "In"
                 "Do" "Loop" "While" "Until" "Select" "Case" "With" "Option" "Explicit" "Lib" "Alias"
                 "True" "False" "Is" "Not" "Nothing" "And" "Or" "On" "Error")
      do (pushnew w vbasense--builtin-keywords :test 'equal))

(defvar vbasense--hash-tli-file-cache (make-hash-table :test 'equal))
(defvar vbasense--hash-tli-guid-cache (make-hash-table :test 'equal))
(defvar vbasense--hash-tli-prog-cache (make-hash-table :test 'equal)) ; key is downcase
(defvar vbasense--hash-app-cache (make-hash-table :test 'equal))      ; key is downcase
(defvar vbasense--hash-object-cache (make-hash-table :test 'equal))   ; key is downcase. have not private module/interface/class
(defvar vbasense--hash-enum-cache (make-hash-table :test 'equal))     ; key is downcase

(defvar vbasense--available-progids nil)
(defvar vbasense--available-applications nil)
(defvar vbasense--available-modules nil)    ; not full name list
(defvar vbasense--available-interfaces nil) ; not full name list
(defvar vbasense--available-classes nil)    ; not full name list
(defvar vbasense--available-coclasses nil)  ; not full name list
(defvar vbasense--available-enums nil)      ; not full name list
(defvar vbasense--available-implicit-enummembers nil)
(defvar vbasense--available-implicit-methods nil)

(defvar vbasense--current-methods nil)
(defvar vbasense--current-classes nil)
(defvar vbasense--current-enums nil)
(defvar vbasense--current-instances nil)
(defvar vbasense--current-instance-alist nil)
(defvar vbasense--current-definition-alist nil)
(make-variable-buffer-local 'vbasense--current-methods)
(make-variable-buffer-local 'vbasense--current-classes)
(make-variable-buffer-local 'vbasense--current-enums)
(make-variable-buffer-local 'vbasense--current-instances)
(make-variable-buffer-local 'vbasense--current-instance-alist)
(make-variable-buffer-local 'vbasense--current-definition-alist)


(defstruct vbasense--app name guid modules interfaces classes enums)
(defstruct vbasense--module name guid (scope 'public) classes enums methods (helptext "") helpfile file)
(defstruct vbasense--interface name guid (scope 'public) methods (helptext "") helpfile file)
(defstruct vbasense--class name guid (scope 'public) coclassp methods (helptext "") helpfile file line)
(defstruct vbasense--enum name guid (scope 'public) members (helptext "") helpfile file line)
(defstruct vbasense--enummember name (helptext "") helpfile file line)
(defstruct vbasense--method name (scope 'public) belongkey (type "") ret arrayretp params (helptext "") helpfile line)
(defstruct vbasense--param name type arrayp ref optional (default ""))


(yaxception:deferror 'vbasense--any-err nil "[VBASense] Failed %s : %s" 'func 'errmsg)


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;; Available for any buffer

(defun vbasense--update-availables ()
  (setq vbasense--available-applications nil)
  (loop for tli being the hash-values in vbasense--hash-app-cache
        do (pushnew (vbasense--app-name tli) vbasense--available-applications :test 'equal))
  (setq vbasense--available-modules nil)
  (setq vbasense--available-interfaces nil)
  (setq vbasense--available-classes nil)
  (setq vbasense--available-coclasses nil)
  (setq vbasense--available-implicit-methods nil)
  (loop for tli being the hash-values in vbasense--hash-object-cache
        do (cond ((vbasense--module-p tli)
                  (pushnew (vbasense--module-name tli) vbasense--available-modules :test 'equal)
                  (loop for m in (vbasense--module-methods tli)
                        do (pushnew (vbasense--method-name m) vbasense--available-implicit-methods :test 'equal)))
                 ((vbasense--interface-p tli)
                  (pushnew (vbasense--interface-name tli) vbasense--available-interfaces :test 'equal))
                 ((vbasense--class-p tli)
                  (pushnew (vbasense--class-name tli) vbasense--available-classes :test 'equal)
                  (when (vbasense--class-coclassp tli)
                    (pushnew (vbasense--class-name tli) vbasense--available-coclasses :test 'equal)
                    (loop for m in (vbasense--class-methods tli)
                          do (pushnew (vbasense--method-name m) vbasense--available-implicit-methods :test 'equal))))))
  (setq vbasense--available-enums nil)
  (setq vbasense--available-implicit-enummembers nil)
  (loop with re-implicit-member = (rx-to-string `(and bos
                                                      (or ,@vbasense-ac-implicit-enummember-prefixes)
                                                      (any "A-Z")))
        for tli being the hash-values in vbasense--hash-enum-cache
        for enumnm = (vbasense--enum-name tli)
        do (pushnew enumnm vbasense--available-enums :test 'equal)
        if (string-match re-implicit-member enumnm)
        do (pushnew enumnm vbasense--available-implicit-enummembers :test 'equal)
        do (loop for m in (vbasense--enum-members tli)
                 for membernm = (vbasense--enummember-name m)
                 if (string-match re-implicit-member membernm)
                 do (pushnew membernm vbasense--available-implicit-enummembers :test 'equal))))


;;;;;;;;;;;;;;;;;;;;;
;; Builtin Library

(yaxception:deferror 'vbasense--ipc-err nil "[VBASense] Failed %s\n%s" 'cmd 'errmsg)
(defvar vbasense--progid-file-name ".progid")

;;;###autoload
(defun vbasense-load-library ()
  "Load library in `vbasense-cache-directory'.

These file are created by getTLI.bat from `vbasense-tli-files' if nesesarry."
  (interactive)
  (yaxception:$
    (yaxception:try
      (vbasense--trace "start load library")
      (when (vbasense--chk-tli-registration)
        (when (not (file-directory-p vbasense-cache-directory))
          (vbasense--debug "mkdir '%s'" vbasense-cache-directory)
          (make-directory vbasense-cache-directory t))
        (vbasense--make-tli-cache)
        (vbasense--make-progid-cache)
        (let ((loaded))
          (dolist (node (directory-files vbasense-cache-directory t))
            (when (and (not (file-directory-p node))
                       (file-exists-p node))
              (let* ((tliid (file-name-nondirectory node)))
                (cond ((and (string= tliid vbasense--progid-file-name)
                            (not vbasense--available-progids))
                       (vbasense--load-progid))
                      ((not (hash-table-p (gethash tliid vbasense--hash-tli-file-cache)))
                       (vbasense--load-tli tliid)
                       (setq loaded t))))))
          (when loaded
            (vbasense--update-availables)))
        (vbasense--trace "finish load library")))
    (yaxception:catch 'vbasense--ipc-err e
      (vbasense--error "failed %s\n%s" (yaxception:get-prop e 'cmd) (yaxception:get-prop e 'errmsg))
      (yaxception:throw e))
    (yaxception:catch 'error e
      (vbasense--error "failed load library : %s\n%s"
                       (yaxception:get-text e) (yaxception:get-stack-trace-string e))
      (yaxception:throw 'vbasense--any-err :func "load library" :errmsg (yaxception:get-text e)))))

;;;###autoload
(defun vbasense-remake-tli-cache (tlifile)
  "Remake library of TLIFILE.

TLIFILE is the path of TypeLibInfo file."
  (interactive
   (list (completing-read "Select file: " vbasense-tli-files nil t)))
  (yaxception:$
    (yaxception:try
      (vbasense--debug "start remake tli cache of %s" tlifile)
      (vbasense--make-tli-cache tlifile t)
      (message "[VBASense] Finished remake tli cache of '%s'." tlifile))
    (yaxception:catch 'error e
      (vbasense--error "failed remake tli cache : %s\n%s"
                       (yaxception:get-text e) (yaxception:get-stack-trace-string e))
      (yaxception:throw 'vbasense--any-err :func "remake tli cache" :errmsg (yaxception:get-text e)))))

;;;###autoload
(defun vbasense-remake-progid-cache ()
  "Remake information of ProgID.

ProgID is a identifier of the library which is installed to Windows."
  (interactive)
  (yaxception:$
    (yaxception:try
      (vbasense--debug "start remake progid cache")
      (vbasense--make-progid-cache t)
      (message "[VBASense] Finished remake progid cache."))
    (yaxception:catch 'error e
      (vbasense--error "failed remake progid cache : %s\n%s"
                       (yaxception:get-text e) (yaxception:get-stack-trace-string e))
      (yaxception:throw 'vbasense--any-err :func "remake progid cache" :errmsg (yaxception:get-text e)))))

;;;###autoload
(defun vbasense-delete-all-cache ()
  "Delete `vbasense-cache-directory'."
  (interactive)
  (yaxception:$
    (yaxception:try
      (when (y-or-n-p "Delete all cache data? (y/n) ")
        (vbasense--debug "start delete all cache")
        (delete-directory vbasense-cache-directory t)
        (setq vbasense--hash-tli-file-cache (make-hash-table :test 'equal))
        (setq vbasense--hash-tli-guid-cache (make-hash-table :test 'equal))
        (setq vbasense--hash-tli-prog-cache (make-hash-table :test 'equal))
        (setq vbasense--hash-app-cache (make-hash-table :test 'equal))
        (setq vbasense--hash-object-cache (make-hash-table :test 'equal))
        (setq vbasense--hash-enum-cache (make-hash-table :test 'equal))
        (vbasense--debug "finish delete all cache")
        (message "[VBASense] Finished delete all cache.")))
    (yaxception:catch 'error e
      (vbasense--error "failed delete all cache : %s\n%s"
                       (yaxception:get-text e) (yaxception:get-stack-trace-string e))
      (yaxception:throw 'vbasense--any-err :func "delete all cache" :errmsg (yaxception:get-text e)))))


(defvar vbasense--own-directory nil)
(defsubst vbasense--get-own-directory ()
  (or vbasense--own-directory
      (setq vbasense--own-directory (loop for dir in load-path
                                          if (directory-files dir nil "\\`vbasense\\.el\\'")
                                          return dir))))

(defvar vbasense--err-buffer-name " *vbasense error*")
(defsubst vbasense--get-err-buffer (&optional init)
  (let ((errbuff (get-buffer-create vbasense--err-buffer-name)))
    (when init
      (with-current-buffer errbuff
        (erase-buffer)))
    errbuff))

(defsubst vbasense--get-err-msg ()
  (with-current-buffer (vbasense--get-err-buffer)
    (buffer-string)))


(defun vbasense--chk-tli-registration ()
  (with-temp-buffer
    (cd (vbasense--get-own-directory))
    (shell-command "cmd.exe /c chkTLI.bat" t (vbasense--get-err-buffer t))
    (let ((errmsg (vbasense--get-err-msg)))
      (cond ((string= errmsg "")
             (vbasense--trace "checked TLBINF32.DLL registration")
             t)
            ((string-match "chkTLI\.vbs\(4, 1\).+TLI\.TLIApplication'?$" errmsg)
             (vbasense--error "not yet regist TLBINF32.DLL")
             (message "[VBASense] Not yet regist TLBINF32.DLL")
             (sleep-for 2)
             nil)
            (t
             (yaxception:throw 'vbasense--ipc-err :cmd "chkTLI.bat" :errmsg errmsg))))))

(defun vbasense--get-tliid (path_or_tliid)
  (let* ((parent (ignore-errors (directory-file-name (file-name-directory path_or_tliid))))
         (tliid (ignore-errors (file-name-sans-extension (file-name-nondirectory path_or_tliid)))))
    (cond ((and (file-exists-p path_or_tliid)
                (not (file-directory-p path_or_tliid)))
           tliid)
          ((and parent
                (file-exists-p parent)
                (not (file-directory-p parent)))
           (concat (file-name-sans-extension (file-name-nondirectory parent))
                   "_"
                   tliid))
          (t
           path_or_tliid))))

(defun vbasense--make-tli-cache (&optional tlifile force)
  (dolist (tlifile (or (when tlifile (list tlifile))
                       vbasense-tli-files))
    (let* ((tliid (vbasense--get-tliid tlifile))
           (cachefile (expand-file-name (concat vbasense-cache-directory "/" tliid)))
           (realpath (replace-regexp-in-string "/[0-9]\\'" "" tlifile)))
      (cond ((not (file-exists-p realpath))
             (vbasense--warn "not found %s" realpath)
             (setq vbasense-tli-files (delete tlifile vbasense-tli-files))
             (message "[VBASense] Not found tli file : %s" realpath)
             (sleep-for 2))
            ((or force
                 (not (file-exists-p cachefile)))
             (vbasense--debug "start make tli of %s" tlifile)
             (with-current-buffer (find-file-noselect cachefile)
               (erase-buffer)
               (cd (vbasense--get-own-directory))
               (setq buffer-file-coding-system 'sjis-dos)
               (message "[VBASense] Getting tli from %s ..." tlifile)
               (let ((coding-system-for-read 'sjis-dos))
                 (shell-command (format "cmd.exe /c getTLI.bat \"%s\"" tlifile) ; not use shell-quote-argument
                                (current-buffer)
                                (vbasense--get-err-buffer t)))
               (save-buffer)
               (kill-buffer))
             (let ((errmsg (vbasense--get-err-msg)))
               (when (not (string= errmsg ""))
                 (delete-file cachefile)
                 (yaxception:throw 'vbasense--ipc-err :cmd "getTLI.bat" :errmsg errmsg)))
             (remhash tliid vbasense--hash-tli-file-cache)
             (vbasense--trace "finish make tli[%s] of %s" tliid tlifile))))))

(defsubst vbasense--get-tli-cache-hash (appnm type name guid)
  (let ((appnm (downcase appnm))
        (name (downcase name)))
    (cond ((string= type "Module")
           (gethash (concat "m:" appnm "." name) vbasense--hash-object-cache))
          ((string= type "Interface")
           (gethash (concat "i:" appnm "." name) vbasense--hash-object-cache))
          ((string= type "CoClass")
           (gethash (concat "c:" appnm "." name) vbasense--hash-object-cache))
          ((string= type "Class")
           (gethash (concat "c:" appnm "." name) vbasense--hash-object-cache))
          ((string= type "Const")
           (gethash (concat appnm "." name) vbasense--hash-enum-cache)))))

(defsubst vbasense--put-tli-cache-hash (appnm type name guid)
  (let ((appnm (downcase appnm))
        (namel (downcase name)))
    (when (string-match "\\`[a-zA-Z]" name)
      (cond ((string= type "Module")
             (let ((ns (make-vbasense--module :name name :guid guid)))
               (puthash (concat "m:" appnm "." namel) ns vbasense--hash-object-cache)
               (puthash (concat "m:" namel) ns vbasense--hash-object-cache)))
            ((string= type "Interface")
             (let ((ns (make-vbasense--interface :name name :guid guid)))
               (puthash (concat "i:" appnm "." namel) ns vbasense--hash-object-cache)
               (puthash (concat "i:" namel) ns vbasense--hash-object-cache)))
            ((string= type "CoClass")
             (let ((ns (make-vbasense--class :name name :guid guid :coclassp t)))
               (puthash (concat "c:" appnm "." namel) ns vbasense--hash-object-cache)
               (puthash (concat "c:" namel) ns vbasense--hash-object-cache)))
            ((string= type "Class")
             (let ((ns (make-vbasense--class :name name :guid guid)))
               (puthash (concat "c:" appnm "." namel) ns vbasense--hash-object-cache)
               (puthash (concat "c:" namel) ns vbasense--hash-object-cache)))
            ((string= type "Const")
             (let ((ns (make-vbasense--enum :name name :guid guid)))
               (puthash (concat appnm "." namel) ns vbasense--hash-enum-cache)
               (puthash namel ns vbasense--hash-enum-cache)))))))

(defsubst vbasense--make-params (members)
  (loop for member in members
        for elements = (split-string member ",")
        for paramnm = (or (pop elements) "")
        for type = (pop elements)
        for optional = (string= (pop elements) "Optional")
        for default = (or (pop elements) "")
        if (not (string= paramnm ""))
        collect (make-vbasense--param :name paramnm :type type :optional optional :default default)))

(defsubst vbasense--make-method (belongnm belongtype members helpstring helpfile)
  (let* ((type (and (listp members) (pop members)))
         (methodnm (and (listp members) (pop members)))
         (ret (and (listp members) (pop members)))
         (params (vbasense--make-params members)))
    (make-vbasense--method :name methodnm
                           :belongkey (concat belongtype ":" (downcase belongnm))
                           :type type
                           :ret ret
                           :params params
                           :helptext helpstring
                           :helpfile helpfile)))

(defsubst vbasense--make-enummember (members helpstring helpfile)
  (let* ((membernm (and (listp members) (pop members))))
    (make-vbasense--enummember :name membernm :helptext helpstring :helpfile helpfile)))

(defun vbasense--load-tli (path_or_tliid)
  (vbasense--debug "start load tli of %s" path_or_tliid)
  (let* ((tliid (vbasense--get-tliid path_or_tliid))
         (fhash (puthash tliid (make-hash-table :test 'equal) vbasense--hash-tli-file-cache))
         app modules interfaces classes enums)
    (with-current-buffer (find-file-noselect (concat vbasense-cache-directory "/" tliid))
      (loop initially (goto-char (point-min))
            until (eobp)
            for line = (replace-regexp-in-string "[\0\r\n]" "" (thing-at-point 'line))
            for parts = (split-string line "###")
            for tlitext = (when (listp parts) (pop parts))
            for helptext = (replace-regexp-in-string "\\\\n" "\n" (if (stringp (car parts)) (pop parts) ""))
            for helpfile = (when (listp parts) (pop parts))
            for tlielem = (when (stringp tlitext) (split-string tlitext "/"))
            for nstype = (or (when (listp tlielem) (pop tlielem))
                             "")
            for nsname = (or (when (listp tlielem) (pop tlielem))
                             "")
            for guid = (when (listp tlielem) (pop tlielem))
            for ns = nil
            for notyet = nil
            do (vbasense--trace "got tli line : %s" line)
            if (string= nstype "App") do (progn (setq ns (setq app (make-vbasense--app :name nsname :guid guid)))
                                                (setq notyet t)
                                                (puthash (downcase nsname) ns vbasense--hash-app-cache))
            else do (when app
                      (setq ns (or (vbasense--get-tli-cache-hash (vbasense--app-name app) nstype nsname guid)
                                   (progn (setq notyet t)
                                          (vbasense--put-tli-cache-hash (vbasense--app-name app) nstype nsname guid))))
                      (cond ((vbasense--module-p ns)
                             (when notyet (push ns modules))
                             (push (vbasense--make-method nsname "m" tlielem helptext helpfile) (vbasense--module-methods ns)))
                            ((vbasense--interface-p ns)
                             (when notyet (push ns interfaces))
                             (push (vbasense--make-method nsname "i" tlielem helptext helpfile) (vbasense--interface-methods ns)))
                            ((vbasense--class-p ns)
                             (when notyet (push ns classes))
                             (push (vbasense--make-method nsname "c" tlielem helptext helpfile) (vbasense--class-methods ns)))
                            ((vbasense--enum-p ns)
                             (when notyet (push ns enums))
                             (push (vbasense--make-enummember tlielem helptext helpfile) (vbasense--enum-members ns)))))
            if (and ns guid notyet) do (puthash guid ns vbasense--hash-tli-guid-cache)
            do (forward-line 1))
      (kill-buffer))
    (when app
      (setf (vbasense--app-modules app) modules)
      (setf (vbasense--app-interfaces app) interfaces)
      (setf (vbasense--app-classes app) classes)
      (setf (vbasense--app-enums app) enums))
    (puthash "app" app fhash)
    (puthash "module" modules fhash)
    (puthash "interface" interfaces fhash)
    (puthash "class" classes fhash)
    (puthash "enum" enums fhash))
  (vbasense--debug "finish load tli of %s" path_or_tliid))


(defun vbasense--make-progid-cache (&optional force)
  (let* ((filepath (concat vbasense-cache-directory "/" vbasense--progid-file-name))
         (buff (when (or force
                         (not (file-exists-p filepath)))
                 (find-file-noselect filepath))))
    (when buff
      (vbasense--debug "start make progid")
      (with-current-buffer buff
        (erase-buffer)
        (cd (vbasense--get-own-directory))
        (setq buffer-file-coding-system 'sjis-dos)
        (message "[VBASense] Getting ProgID ...")
        (let ((coding-system-for-read 'sjis-dos))
          (shell-command "cmd.exe /c getProgIDAll.bat" (current-buffer) (vbasense--get-err-buffer t)))
        (save-buffer)
        (kill-buffer))
      (let ((errmsg (vbasense--get-err-msg)))
        (when (not (string= errmsg ""))
          (delete-file filepath)
          (yaxception:throw 'vbasense--ipc-err :cmd "getProgIDAll.bat" :errmsg errmsg)))
      (setq vbasense--available-progids nil)
      (vbasense--trace "finish make progid"))))

(defun vbasense--load-progid ()
  (vbasense--debug "start load progid")
  (with-current-buffer (find-file-noselect (concat vbasense-cache-directory "/" vbasense--progid-file-name))
    (loop initially (goto-char (point-min))
          for line = (replace-regexp-in-string "[\0\r\n]" "" (thing-at-point 'line))
          for infolist = (split-string line "/")
          for type = (when (listp infolist) (pop infolist))
          for progid = (when (listp infolist) (pop infolist))
          for guid = (when (listp infolist) (pop infolist))
          until (eobp)
          if (and (stringp type)
                  (string= type "ProgID"))
          do (progn (pushnew progid vbasense--available-progids :test 'equal)
                    (puthash (downcase progid) guid vbasense--hash-tli-prog-cache))
          do (forward-line 1))
    (kill-buffer))
  (vbasense--debug "finish load progid"))


;;;;;;;;;;;;;;;;;;
;; User Library

(defsubst vbasense--vba-file-p (file)
  (and file
       (file-exists-p file)
       (not (file-directory-p file))
       (string-match "\\.\\(cls\\|bas\\|frm\\)\\'" file)))

(defsubst vbasense--active-p ()
  (memq major-mode vbasense-enable-modes))

;;;###autoload
(defun* vbasense-load-file (file &key delay not-update-cache)
  "Load user file."
  (interactive (list (read-file-name "Load file: " t nil t)))
  (yaxception:$
    (yaxception:try
      (vbasense--debug "start load file. file[%s] delay[%s] not-update-cache[%s]" file delay not-update-cache)
      (let* ((tliid (vbasense--get-tliid file))
             (timernm (concat "vbasense--timer-load-" tliid)))
        (when (and (vbasense--vba-file-p file)
                   (not (gethash tliid vbasense--hash-tli-file-cache))
                   (or (not (intern-soft timernm))
                       (not (timerp (symbol-value (intern timernm))))))
          (cond (delay
                 (set (intern timernm)
                      (run-with-timer (vbasense--get-timer-delays)
                                      nil
                                      (lambda (file not-update-cache)
                                        (vbasense--analyze-file file)
                                        (when (not not-update-cache)
                                          (vbasense--update-availables)))
                                      file
                                      not-update-cache)))
                (t
                 (vbasense--analyze-file file)
                 (when (not not-update-cache)
                   (vbasense--update-availables))))))
      (vbasense--debug "finish load file. file[%s]" file))
    (yaxception:catch 'error e
      (vbasense--error "failed load file : %s\n%s"
                       (yaxception:get-text e) (yaxception:get-stack-trace-string e))
      (yaxception:throw 'vbasense--any-err :func "load file" :errmsg (yaxception:get-text e)))))

;;;###autoload
(defun vbasense-load-directory (dir)
  "Load user file in directory."
  (interactive (list (read-directory-name "Load file in directory: " t nil t default-directory)))
  (yaxception:$
    (yaxception:try
      (vbasense--debug "start load directory : %s" dir)
      (vbasense--load-files-with-timer (eshell-extended-glob (concat (file-name-as-directory dir) "**/*")))
      (vbasense--debug "finish load directory : %s" dir))
    (yaxception:catch 'error e
      (vbasense--error "failed load directory : %s\n%s"
                       (yaxception:get-text e) (yaxception:get-stack-trace-string e))
      (yaxception:throw 'vbasense--any-err :func "load directory" :errmsg (yaxception:get-text e)))))

;;;###autoload
(defun vbasense-load-recently ()
  "Load user file in recentf."
  (interactive)
  (yaxception:$
    (yaxception:try
      (vbasense--debug "start load recently.")
      (vbasense--load-files-with-timer recentf-list)
      (vbasense--debug "finish load recently."))
    (yaxception:catch 'error e
      (vbasense--error "failed load recently : %s\n%s"
                       (yaxception:get-text e) (yaxception:get-stack-trace-string e))
      (yaxception:throw 'vbasense--any-err :func "load recently" :errmsg (yaxception:get-text e)))))

;;;###autoload
(defun vbasense-load-project ()
  "Load user file in project.

The project is detected by `anything-project'."
  (interactive)
  (yaxception:$
    (yaxception:try
      (vbasense--debug "start load project.")
      (when (and (featurep 'anything-project)
                 (functionp 'ap:get-project-files)
                 (functionp 'ap:expand-file))
        (vbasense--load-files-with-timer (loop for f in (ap:get-project-files)
                                               collect (ap:expand-file f))))
      (vbasense--debug "finish load project."))
    (yaxception:catch 'error e
      (vbasense--error "failed load project : %s\n%s"
                       (yaxception:get-text e) (yaxception:get-stack-trace-string e))
      (yaxception:throw 'vbasense--any-err :func "load project" :errmsg (yaxception:get-text e)))))

;;;###autoload
(defun vbasense-reload-current-buffer ()
  "Reload user file of current buffer."
  (interactive)
  (yaxception:$
    (yaxception:try
      (when (and (vbasense--vba-file-p (buffer-file-name))
                 (vbasense--active-p))
        (vbasense--debug "start reload current buffer.")
        (vbasense--analyze-file (buffer-file-name))
        (vbasense--update-availables)
        (vbasense--update-current-definition)
        (vbasense--debug "finish reload current buffer.")))
    (yaxception:catch 'error e
      (vbasense--error "failed reload current buffer : %s\n%s"
                       (yaxception:get-text e) (yaxception:get-stack-trace-string e))
      (yaxception:throw 'vbasense--any-err :func "reload current buffer" :errmsg (yaxception:get-text e)))))

(defvar vbasense--running-job-interval 2)
(defvar vbasense--last-scheduling-job-time 0)
(defun vbasense--get-timer-delays ()
  (let* ((now (car (current-time)))
         (until-last (cond ((> vbasense--last-scheduling-job-time now)
                            (- vbasense--last-scheduling-job-time now))
                           (t
                            0))))
    (setq vbasense--last-scheduling-job-time (+ now until-last vbasense--running-job-interval))
    (+ until-last vbasense--running-job-interval)))

(defun vbasense--load-files-with-timer (files)
  (when (listp files)
    (loop with flist = (loop for f in files
                             if (vbasense--vba-file-p f)
                             collect f)
          while flist
          for f = (pop flist)
          for not-update-cache = (when flist t)
          do (vbasense-load-file f :delay t :not-update-cache not-update-cache))))


;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;; Analyze Buffer Content

(defsubst vbasense--get-user-tli-type (fpath)
  (let* ((tliid (vbasense--get-tliid fpath))
         (suffix (or (file-name-extension fpath)
                     ""))
         (prefix (or (ignore-errors (substring (file-name-nondirectory fpath) 0 3))
                     "")))
    (cond ((string= suffix "bas")
           'module)
          ((and (string= suffix "cls")
                (string-match "^I[A-Z]" prefix))
           'interface)
          ((or (string= suffix "cls")
               (string= suffix "frm"))
           'class))))

(defsubst vbasense--new-user-tli (fpath)
  (let* ((tliid (vbasense--get-tliid fpath))
         (tlitype (vbasense--get-user-tli-type fpath)))
    (case tlitype
      (module    (make-vbasense--module :name tliid :helpfile fpath :file fpath))
      (interface (make-vbasense--interface :name tliid :helpfile fpath :file fpath))
      (class     (make-vbasense--class :name tliid :helpfile fpath :file fpath)))))

(defsubst vbasense--get-object-hash-key (tli)
  (cond ((vbasense--module-p tli)    (concat "m:" (downcase (vbasense--module-name tli))))
        ((vbasense--interface-p tli) (concat "i:" (downcase (vbasense--interface-name tli))))
        ((vbasense--class-p tli)     (concat "c:" (downcase (vbasense--class-name tli))))
        (t                           "")))

(defsubst vbasense--get-object-hash-key-of-current ()
  (let* ((tliid (vbasense--get-tliid (buffer-file-name)))
         (tlitype (vbasense--get-user-tli-type (buffer-file-name)))
         (typech (case tlitype
                   (module    "m")
                   (interface "i")
                   (class     "c"))))
    (if (not typech)
        ""
      (concat typech ":" (downcase tliid)))))

(defun vbasense--analyze-file (fpath)
  (with-temp-buffer
    (insert-file-contents fpath)
    (let* ((tliid (vbasense--get-tliid fpath))
           (fhash (puthash tliid (make-hash-table :test 'equal) vbasense--hash-tli-file-cache))
           (tli (vbasense--new-user-tli fpath))
           (belongkey (vbasense--get-object-hash-key tli))
           (methods (when (not (string= belongkey ""))
                      (loop for m in (append (vbasense--get-functions-in-buffer   belongkey :buff-file fpath :exclude-private t)
                                             (vbasense--get-subs-in-buffer        belongkey :buff-file fpath :exclude-private t)
                                             (vbasense--get-members-in-buffer     belongkey :buff-file fpath :exclude-private t)
                                             (vbasense--get-propgetters-in-buffer belongkey :buff-file fpath :exclude-private t)
                                             (vbasense--get-propsetters-in-buffer belongkey :buff-file fpath :exclude-private t))
                            if (vbasense--method-p m) collect m)))
           (enums (vbasense--get-enums-in-buffer :buff-file fpath))
           (classes (loop for c in `(,tli ,@(vbasense--get-types-in-buffer :buff-file fpath))
                          if (vbasense--class-p c) collect c))
           (interfaces (when (vbasense--interface-p tli) (list tli)))
           (modules (when (vbasense--module-p tli) (list tli))))
      ;; regist method to tli
      (cond ((vbasense--class-p tli)     (setf (vbasense--class-methods tli) methods))
            ((vbasense--interface-p tli) (setf (vbasense--interface-methods tli) methods))
            ((vbasense--module-p tli)    (setf (vbasense--module-methods tli) methods)))
      ;; build tli file hash
      (puthash "app" nil fhash)
      (puthash "module" modules fhash)
      (puthash "interface" interfaces fhash)
      (puthash "class" classes fhash)
      (puthash "enum" enums fhash)
      ;; regist tli to global hash
      (cond ((string= belongkey "c:thisworkbook")
             nil)
            ((not (string= belongkey ""))
             (puthash belongkey tli vbasense--hash-object-cache)))
      ;; regist current class/enum to global hash
      (let ((mdlnm (when (and (vbasense--module-p tli)
                              (not (eq (vbasense--module-scope tli) 'private)))
                     (downcase (vbasense--module-name tli)))))
        (loop for c in classes
              for classnm = (downcase (vbasense--class-name c))
              if (not (eq (vbasense--class-scope c) 'private))
              do (progn (puthash (concat "c:" classnm) c vbasense--hash-object-cache)
                        (when mdlnm
                          (push c (vbasense--module-classes tli))
                          (puthash (concat "c:" mdlnm "." classnm) c vbasense--hash-object-cache))))
        (loop for e in enums
              for enumnm = (downcase (vbasense--enum-name e))
              if (not (eq (vbasense--enum-scope e) 'private))
              do (progn (puthash enumnm e vbasense--hash-enum-cache)
                        (when mdlnm
                          (push e (vbasense--module-enums tli))
                          (puthash (concat mdlnm "." enumnm) e vbasense--hash-enum-cache))))))))

(defun vbasense--update-current-definition ()
  (setq vbasense--current-definition-alist nil)
  (let ((belongkey (vbasense--get-object-hash-key-of-current)))
    (setq vbasense--current-methods (when (not (string= belongkey ""))
                                      (loop for m in (append (vbasense--get-functions-in-buffer   belongkey :return-name t)
                                                             (vbasense--get-subs-in-buffer        belongkey :return-name t)
                                                             (vbasense--get-propgetters-in-buffer belongkey :return-name t)
                                                             (vbasense--get-propsetters-in-buffer belongkey :return-name t))
                                            if m collect m)))
    (setq vbasense--current-classes (vbasense--get-types-in-buffer :return-name t))
    (setq vbasense--current-enums (vbasense--get-enums-in-buffer :return-name t))
    (vbasense--update-current-instance)))

(defsubst vbasense--get-current-definition (defnm deftype)
  (let* ((keytype (case deftype
                    ((function sub) 'method)
                    (t              deftype)))
         (key (concat (symbol-name keytype) ":" defnm)))
    (or (assoc-default key vbasense--current-definition-alist)
        (let* ((belongkey (vbasense--get-object-hash-key-of-current))
               (def (case deftype
                      (function (loop for fnc in '(vbasense--get-members-in-buffer
                                                   vbasense--get-functions-in-buffer
                                                   vbasense--get-propgetters-in-buffer)
                                      for f = (loop for f in (funcall fnc belongkey :specified-name defnm)
                                                    for ftype = (vbasense--method-type f)
                                                    if (or (string= ftype "Function")
                                                           (string= ftype "Property Get"))
                                                    return f)
                                      if f return f))
                      (sub (loop for fnc in '(vbasense--get-members-in-buffer
                                              vbasense--get-subs-in-buffer
                                              vbasense--get-propsetters-in-buffer)
                                 for f = (loop for f in (funcall fnc belongkey :specified-name defnm)
                                               for ftype = (vbasense--method-type f)
                                               if (or (string= ftype "Sub")
                                                      (string= ftype "Property Let"))
                                               return f)
                                 if f return f))
                      (method (loop for fnc in '(vbasense--get-members-in-buffer
                                                 vbasense--get-functions-in-buffer
                                                 vbasense--get-subs-in-buffer
                                                 vbasense--get-propgetters-in-buffer
                                                 vbasense--get-propsetters-in-buffer)
                                    for methods = (funcall fnc belongkey :specified-name defnm)
                                    if methods return (pop methods)))
                      (class (let ((classes (vbasense--get-types-in-buffer :specified-name defnm)))
                               (when classes (pop classes))))
                      (enum (let ((enums (vbasense--get-enums-in-buffer :specified-name defnm)))
                              (when enums (pop enums)))))))
          (when def
            (push `(,key . ,def) vbasense--current-definition-alist)
            def)))))

(defsubst vbasense--get-method-of-tli (tli methodnm &optional methodtype)
  (loop with methodnm = (downcase methodnm)
        for m in (cond ((vbasense--module-p tli)    (vbasense--module-methods tli))
                       ((vbasense--interface-p tli) (vbasense--interface-methods tli))
                       ((vbasense--class-p tli)     (vbasense--class-methods tli)))
        for currnm = (downcase (vbasense--method-name m))
        for currtype = (vbasense--method-type m)
        if (and (string= currnm methodnm)
                (or (not methodtype)
                    (and (eq methodtype 'function)
                         (or (string= currtype "Function")
                             (string= currtype "Property Get")))
                    (and (eq methodtype 'sub)
                         (or (string= currtype "Sub")
                             (string= currtype "Property Let")))))
        return m))

(defsubst vbasense--get-latest-class-definition (classnm)
  (or (when (member classnm vbasense--current-classes)
        (vbasense--get-current-definition classnm 'class))
      (gethash (concat "c:" (downcase classnm)) vbasense--hash-object-cache)))

(defsubst vbasense--get-latest-enum-definition (enumnm)
  (or (when (member enumnm vbasense--current-enums)
        (vbasense--get-current-definition enumnm 'enum))
      (gethash (downcase enumnm) vbasense--hash-enum-cache)))

(defsubst vbasense--get-latest-method-definition (nsnm nstype methodnm &optional methodtype)
  (let ((tli (case nstype
               (module    (gethash (concat "m:" (downcase nsnm)) vbasense--hash-object-cache))
               (interface (gethash (concat "i:" (downcase nsnm)) vbasense--hash-object-cache))
               (class     (vbasense--get-latest-class-definition nsnm)))))
    (when tli
      (cond ((string= (vbasense--get-object-hash-key tli)
                      (vbasense--get-object-hash-key-of-current))
             (if methodtype
                 (vbasense--get-current-definition methodnm methodtype)
               (vbasense--get-current-definition methodnm 'method)))
            (t
             (vbasense--get-method-of-tli tli methodnm methodtype))))))


(defvar vbasense--regexp-param-name (rx-to-string `(and bos
                                                        (regexp ,vbasense--regexp-variable-ident)
                                                        (group (? "()"))
                                                        eos)))
(defsubst vbasense--get-params-in-string (str)
  (loop for paramstr in (split-string str ",")
        for p = (loop with paramnm = ""
                      with paramtype = ""
                      with is-ref = t
                      with defvalue = ""
                      with is-optional = nil
                      with next-is-type = nil
                      with next-is-default = nil
                      with arrayp = nil
                      for e in (split-string paramstr " +")
                      for le = (downcase e)
                      do (cond ((string= le "byval")
                                (setq is-ref nil))
                               ((string= le "byref")
                                (setq is-ref t))
                               ((string= le "optional")
                                (setq is-optional t))
                               ((string= le "paramarray")
                                (setq arrayp t))
                               ((string= le "as")
                                (setq next-is-type t))
                               ((string= e "=")
                                (setq next-is-default t))
                               ((and next-is-type
                                     (not (string= e "")))
                                (setq paramtype e)
                                (setq next-is-type nil))
                               ((and is-optional
                                     next-is-default
                                     (not (string= e "")))
                                (setq defvalue e)
                                (setq next-is-default nil))
                               ((and (string= paramnm "")
                                     (not next-is-type)
                                     (not next-is-default)
                                     (string-match vbasense--regexp-param-name e))
                                (setq paramnm e)
                                (when (not (string= (match-string-no-properties 1 e) ""))
                                  (setq arrayp t))))
                      finally return (when (not (string= paramnm ""))
                                       (make-vbasense--param :name paramnm
                                                             :type paramtype
                                                             :arrayp arrayp
                                                             :ref is-ref
                                                             :optional is-optional
                                                             :default defvalue)))
        if (vbasense--param-p p) collect p))

(defsubst vbasense--get-scope (str)
  (if (or (not (stringp str))
          (string= str ""))
      'public
    (intern (downcase (replace-regexp-in-string "\\s-+\\'" "" str)))))

(defsubst vbasense--get-line-number (pt)
  (+ (count-lines (point-min) pt) 1))

(defsubst* vbasense--get-region-endpt (region-names &key return-maxpt-if-not-found)
  (save-excursion
    (let ((currpt (point))
          (re (when (listp region-names)
                (rx-to-string `(and bol (* space)
                                    (or ,@vbasense--words-end) (+ space)
                                    (or ,@region-names))))))
      (cond ((and re (re-search-forward re nil t)) (point))
            (return-maxpt-if-not-found             (point-max))
            (t                                     currpt)))))

(defvar vbasense--regexp-user-member (rx-to-string `(and bol (* space)
                                                         (group (or ,@vbasense--words-scope) (+ space))
                                                         (? (or ,@vbasense--words-static
                                                                ,@vbasense--words-const) (+ space))
                                                         (group (regexp ,vbasense--regexp-variable-ident))
                                                         (group (? "(" (* (not (any ")"))) ")")) (+ space)
                                                         (or ,@vbasense--words-as) (+ space)
                                                         (? (or ,@vbasense--words-new) (+ space))
                                                         (group (regexp ,vbasense--regexp-namespace-ident)))))
(defun* vbasense--get-members-in-buffer (belongkey &key exclude-private return-name specified-name buff-file)
  (save-excursion
    (loop with fpath = (or buff-file (buffer-file-name))
          initially (goto-char (point-min))
          while (re-search-forward vbasense--regexp-user-member nil t)
          for scope = (vbasense--get-scope (match-string-no-properties 1))
          for name = (match-string-no-properties 2)
          for arrayretp = (not (string= (match-string-no-properties 3) ""))
          for ret = (match-string-no-properties 4)
          for linenumber = (vbasense--get-line-number (match-beginning 1))
          for params = (when (not return-name)
                         (list (make-vbasense--param :name (downcase name)
                                                     :type ret
                                                     :arrayp arrayretp
                                                     :ref arrayretp)))
          if (and (or (not exclude-private)
                      (not (eq scope 'private)))
                  (or (not specified-name)
                      (string= specified-name name)))
          append (if return-name
                     (list name)
                   (list (make-vbasense--method :name name
                                                :belongkey belongkey
                                                :scope scope
                                                :type "Property Let"
                                                :params params
                                                :helpfile fpath
                                                :line linenumber)
                         (make-vbasense--method :name name
                                                :belongkey belongkey
                                                :scope scope
                                                :type "Property Get"
                                                :ret ret
                                                :arrayretp arrayretp
                                                :helpfile fpath
                                                :line linenumber))))))

(defvar vbasense--regexp-user-function (rx-to-string `(and bol (* space)
                                                           (group (? (or ,@vbasense--words-scope) (+ space)))
                                                           (? (or ,@vbasense--words-static) (+ space))
                                                           (or ,@vbasense--words-function) (+ space)
                                                           (group (regexp ,vbasense--regexp-variable-ident))
                                                           (regexp ,vbasense--regexp-argument) (+ space)
                                                           (or ,@vbasense--words-as) (+ space)
                                                           (group (regexp ,vbasense--regexp-namespace-ident)))))
(defun* vbasense--get-functions-in-buffer (belongkey &key exclude-private return-name specified-name buff-file)
  (save-excursion
    (loop with fpath = (or buff-file (buffer-file-name))
          initially (goto-char (point-min))
          while (re-search-forward vbasense--regexp-user-function nil t)
          for scope = (vbasense--get-scope (match-string-no-properties 1))
          for name = (match-string-no-properties 2)
          for linenumber = (vbasense--get-line-number (match-beginning 1))
          for ret = (match-string-no-properties 4)
          for params = (when (not return-name)
                         (vbasense--get-params-in-string (match-string-no-properties 3)))
          if (and (or (not exclude-private)
                      (not (eq scope 'private)))
                  (or (not specified-name)
                      (string= specified-name name)))
          collect (if return-name
                      name
                    (make-vbasense--method :name name
                                           :belongkey belongkey
                                           :scope scope
                                           :type "Function"
                                           :ret ret
                                           :params params
                                           :helpfile fpath
                                           :line linenumber)))))

(defvar vbasense--regexp-user-sub (rx-to-string `(and bol (* space)
                                                      (group (? (or ,@vbasense--words-scope) (+ space)))
                                                      (? (or ,@vbasense--words-static) (+ space))
                                                      (or ,@vbasense--words-sub) (+ space)
                                                      (group (regexp ,vbasense--regexp-variable-ident))
                                                      (regexp ,vbasense--regexp-argument))))
(defun* vbasense--get-subs-in-buffer (belongkey &key exclude-private return-name specified-name buff-file)
  (save-excursion
    (loop with fpath = (or buff-file (buffer-file-name))
          initially (goto-char (point-min))
          while (re-search-forward vbasense--regexp-user-sub nil t)
          for scope = (vbasense--get-scope (match-string-no-properties 1))
          for name = (match-string-no-properties 2)
          for linenumber = (vbasense--get-line-number (match-beginning 1))
          for params = (when (not return-name)
                         (vbasense--get-params-in-string (match-string-no-properties 3)))
          if (and (or (not exclude-private)
                      (not (eq scope 'private)))
                  (or (not specified-name)
                      (string= specified-name name)))
          collect (if return-name
                      name
                    (make-vbasense--method :name name
                                           :belongkey belongkey
                                           :scope scope
                                           :type "Sub"
                                           :params params
                                           :helpfile fpath
                                           :line linenumber)))))

(defvar vbasense--regexp-user-propgetter (rx-to-string `(and bol (* space)
                                                             (group (? (or ,@vbasense--words-scope) (+ space)))
                                                             (? (or ,@vbasense--words-static) (+ space))
                                                             (or ,@vbasense--words-property) (+ space)
                                                             (or ,@vbasense--words-get) (+ space)
                                                             (group (regexp ,vbasense--regexp-variable-ident))
                                                             (regexp ,vbasense--regexp-argument) (+ space)
                                                             (or ,@vbasense--words-as) (+ space)
                                                             (group (regexp ,vbasense--regexp-namespace-ident)))))
(defun* vbasense--get-propgetters-in-buffer (belongkey &key exclude-private return-name specified-name buff-file)
  (save-excursion
    (loop with fpath = (or buff-file (buffer-file-name))
          initially (goto-char (point-min))
          while (re-search-forward vbasense--regexp-user-propgetter nil t)
          for scope = (vbasense--get-scope (match-string-no-properties 1))
          for name = (match-string-no-properties 2)
          for linenumber = (vbasense--get-line-number (match-beginning 1))
          for params = (when (not return-name)
                         (vbasense--get-params-in-string (match-string-no-properties 3)))
          for ret = (match-string-no-properties 4)
          if (and (or (not exclude-private)
                      (not (eq scope 'private)))
                  (or (not specified-name)
                      (string= specified-name name)))
          collect (if return-name
                      name
                    (make-vbasense--method :name name
                                           :belongkey belongkey
                                           :scope scope
                                           :type "Property Get"
                                           :params params
                                           :ret ret
                                           :helpfile fpath
                                           :line linenumber)))))

(defvar vbasense--regexp-user-propsetter (rx-to-string `(and bol (* space)
                                                             (group (? (or ,@vbasense--words-scope) (+ space)))
                                                             (? (or ,@vbasense--words-static) (+ space))
                                                             (or ,@vbasense--words-property) (+ space)
                                                             (or ,@vbasense--words-let) (+ space)
                                                             (group (regexp ,vbasense--regexp-variable-ident))
                                                             (regexp ,vbasense--regexp-argument))))
(defun* vbasense--get-propsetters-in-buffer (belongkey &key exclude-private return-name specified-name buff-file)
  (save-excursion
    (loop with fpath = (or buff-file (buffer-file-name))
          initially (goto-char (point-min))
          while (re-search-forward vbasense--regexp-user-propsetter nil t)
          for scope = (vbasense--get-scope (match-string-no-properties 1))
          for name = (match-string-no-properties 2)
          for linenumber = (vbasense--get-line-number (match-beginning 1))
          for params = (when (not return-name)
                         (vbasense--get-params-in-string (match-string-no-properties 3)))
          if (and (or (not exclude-private)
                      (not (eq scope 'private)))
                  (or (not specified-name)
                      (string= specified-name name)))
          collect (if return-name
                      name
                    (make-vbasense--method :name name
                                           :belongkey belongkey
                                           :scope scope
                                           :type "Property Let"
                                           :params params
                                           :helpfile fpath
                                           :line linenumber)))))

(defvar vbasense--regexp-user-type-line (rx-to-string `(and bol (* space)
                                                            (group (regexp ,vbasense--regexp-variable-ident)) (+ space)
                                                            (or ,@vbasense--words-as) (+ space)
                                                            (group (regexp ,vbasense--regexp-namespace-ident)))))
(defsubst vbasense--get-type-methods-in-string (str belongkey fpath linenumber)
  (loop for line in (split-string str "\r?\n")
        if (string-match vbasense--regexp-user-type-line line)
        append (let* ((name (match-string-no-properties 1 line))
                      (ret (match-string-no-properties 2 line))
                      (params (vbasense--get-params-in-string (concat "hoge As " ret))))
                 (list (make-vbasense--method :name name
                                              :belongkey belongkey
                                              :type "Property Let"
                                              :params params
                                              :helpfile fpath
                                              :line linenumber)
                       (make-vbasense--method :name name
                                              :belongkey belongkey
                                              :type "Property Get"
                                              :ret ret
                                              :helpfile fpath
                                              :line linenumber)))))

(defvar vbasense--regexp-user-type (rx-to-string `(and bol (* space)
                                                       (group (? (or ,@vbasense--words-scope) (+ space)))
                                                       (or ,@vbasense--words-type) (+ space)
                                                       (group (regexp ,vbasense--regexp-variable-ident)))))

(defun* vbasense--get-types-in-buffer (&key exclude-private return-name specified-name buff-file)
  (save-excursion
    (loop with fpath = (or buff-file (buffer-file-name))
          initially (goto-char (point-min))
          while (re-search-forward vbasense--regexp-user-type nil t)
          for scope = (vbasense--get-scope (match-string-no-properties 1))
          for name = (match-string-no-properties 2)
          for linenumber = (vbasense--get-line-number (match-beginning 1))
          if (and (or (not exclude-private)
                      (not (eq scope 'private)))
                  (or (not specified-name)
                      (string= specified-name name)))
          collect (if return-name
                      name
                    (make-vbasense--class :name name
                                          :scope scope
                                          :methods (vbasense--get-type-methods-in-string
                                                    (buffer-substring-no-properties (point)
                                                                                    (vbasense--get-region-endpt vbasense--words-type))
                                                    (concat "c:" name)
                                                    fpath
                                                    linenumber)
                                          :helpfile fpath
                                          :file fpath
                                          :line linenumber)))))

(defvar vbasense--regexp-enum-member (rx-to-string `(and bol (* space)
                                                         (group (regexp ,vbasense--regexp-variable-ident)) (+ space)
                                                         "=" (+ space) (+ (any "0-9")))))
(defsubst vbasense--get-enum-members-in-string (str file line)
  (loop for code in (split-string str "\r?\n")
        for membernm = (when (string-match vbasense--regexp-enum-member code)
                         (match-string-no-properties 1 code))
        if membernm
        collect (make-vbasense--enummember :name membernm :file file :line line)))

(defvar vbasense--regexp-user-enum (rx-to-string `(and bol (* space)
                                                       (group (? (or ,@vbasense--words-scope) (+ space)))
                                                       (or ,@vbasense--words-enum) (+ space)
                                                       (group (regexp ,vbasense--regexp-variable-ident)))))

(defun* vbasense--get-enums-in-buffer (&key exclude-private return-name specified-name buff-file)
  (save-excursion
    (loop with fpath = (or buff-file (buffer-file-name))
          initially (goto-char (point-min))
          while (re-search-forward vbasense--regexp-user-enum nil t)
          for scope = (vbasense--get-scope (match-string-no-properties 1))
          for name = (match-string-no-properties 2)
          for linenumber = (vbasense--get-line-number (match-beginning 1))
          if (and (or (not exclude-private)
                      (not (eq scope 'private)))
                  (or (not specified-name)
                      (string= specified-name name)))
          collect (if return-name
                      name
                    (let* ((endpt (vbasense--get-region-endpt vbasense--words-enum))
                           (mtext (buffer-substring-no-properties (point) endpt)))
                      (make-vbasense--enum :name name
                                           :scope scope
                                           :members (vbasense--get-enum-members-in-string mtext fpath linenumber)
                                           :helpfile fpath
                                           :file fpath
                                           :line linenumber))))))

(defvar vbasense--regexp-definition-start (rx-to-string `(and bol (* space)
                                                              (? (or ,@vbasense--words-scope) (+ space))
                                                              (? (or ,@vbasense--words-static) (+ space))
                                                              (group (or ,@vbasense--words-function
                                                                         ,@vbasense--words-sub
                                                                         ,@vbasense--words-property
                                                                         ,@vbasense--words-type
                                                                         ,@vbasense--words-enum))
                                                              (+ space)
                                                              (? (or ,@vbasense--words-get
                                                                     ,@vbasense--words-let) (+ space))
                                                              (group (regexp ,vbasense--regexp-variable-ident)))))
(defvar vbasense--regexp-argument-start (rx-to-string `(and point (regexp ,vbasense--regexp-argument))))
(defvar vbasense--regexp-current-variable (rx-to-string `(and bol (* space)
                                                              (or ,@vbasense--words-dim
                                                                  ,@vbasense--words-static
                                                                  ,@vbasense--words-const)
                                                              (+ space)
                                                              (group (regexp ,vbasense--regexp-variable-ident))
                                                              (? "(" (* (not (any ")"))) ")") (+ space)
                                                              (or ,@vbasense--words-as) (+ space)
                                                              (? (or ,@vbasense--words-new) (+ space))
                                                              (group (regexp ,vbasense--regexp-namespace-ident)))))
(defvar vbasense--last-region-text "")
(defun vbasense--update-current-instance ()
  (save-excursion
    (let ((currpt (point)))
      (vbasense--trace "start update current instance")
      (if (not (re-search-backward vbasense--regexp-definition-start nil t))
          (progn (setq vbasense--current-instances nil)
                 (setq vbasense--current-instance-alist nil)
                 (setq vbasense--last-region-text ""))
        (vbasense--trace "start get current instance region")
        (let* ((defnm (match-string-no-properties 1))
               (defident (match-string-no-properties 2))
               (defstartpt (match-end 0))
               (defendpt (vbasense--get-region-endpt (list defnm) :return-maxpt-if-not-found t))
               (deftext (buffer-substring-no-properties defstartpt defendpt)))
          (vbasense--trace "got current instance region : point[%s] type[%s] name[%s] start[%s] end[%s]"
                           currpt defnm defident defstartpt defendpt)
          (when (and (> currpt defstartpt)
                     (< currpt defendpt)
                     (not (string= deftext vbasense--last-region-text)))
            (vbasense--trace "start find current instance")
            (setq vbasense--current-instances nil)
            (setq vbasense--current-instance-alist nil)
            (setq vbasense--last-region-text deftext)
            ;; find member
            (loop initially (goto-char (point-min))
                  while (re-search-forward vbasense--regexp-user-member nil t)
                  for varnm = (match-string-no-properties 2)
                  for vartype = (match-string-no-properties 4)
                  do (progn (push `(,varnm . ,vartype) vbasense--current-instance-alist)
                            (pushnew varnm vbasense--current-instances :test 'equal)))
            ;; find argument
            (goto-char defstartpt)
            (when (re-search-forward vbasense--regexp-argument-start nil t)
              (loop with arglastpt = (match-end 0)
                    for p in (vbasense--get-params-in-string (match-string-no-properties 1))
                    for varnm = (vbasense--param-name p)
                    for vartype = (vbasense--param-type p)
                    if (and (not (string= varnm ""))
                            (not (string= vartype ""))
                            (not (assoc-default varnm vbasense--current-instance-alist)))
                    do (progn (push `(,varnm . ,vartype) vbasense--current-instance-alist)
                              (pushnew varnm vbasense--current-instances :test 'equal))
                    finally do (goto-char (setq defstartpt arglastpt))))
            ;; find variable
            (forward-line)
            (while (re-search-forward vbasense--regexp-current-variable defendpt t)
              (let ((varnm (match-string-no-properties 1))
                    (vartype (match-string-no-properties 2)))
                (when (not (assoc-default varnm vbasense--current-instance-alist))
                  (push `(,varnm . ,vartype) vbasense--current-instance-alist)
                  (pushnew varnm vbasense--current-instances :test 'equal))))
            (vbasense--embody-instance defstartpt defendpt)
            (vbasense--trace "found current instance : %s" vbasense--current-instance-alist)))))))

(defun vbasense--embody-instance (startpt endpt)
  (let ((srchvars (loop for e in vbasense--current-instance-alist
                        for varnm = (car e)
                        for vartype = (downcase (cdr e))
                        if (string= vartype "object")
                        collect varnm)))
    ;; find createobject statement
    (when srchvars
      (goto-char (point-min))
      (let ((re (rx-to-string `(and bol (* space)
                                    (or ,@vbasense--words-set) (+ space)
                                    (group (or ,@srchvars)) (+ space) "=" (+ space)
                                    (or ,@vbasense--words-createobject)
                                    "(" (* space) "\"" (group (+ (any "a-zA-Z0-9_.")))))))
        (while (re-search-forward re nil t)
          (let* ((varnm (match-string-no-properties 1))
                 (progid (match-string-no-properties 2))
                 (currdef (assoc varnm vbasense--current-instance-alist))
                 (guid (gethash (downcase progid) vbasense--hash-tli-prog-cache))
                 (tli (when guid (gethash guid vbasense--hash-tli-guid-cache)))
                 (vartype (when tli (cond ((vbasense--module-p tli)    (vbasense--module-name tli))
                                          ((vbasense--interface-p tli) (vbasense--interface-name tli))
                                          ((vbasense--class-p tli)     (vbasense--class-name tli))
                                          ((vbasense--enum-p tli)      (vbasense--enum-name tli))))))
            (when (not guid) (vbasense--info "can't find GUID from ProgID[%s]" progid))
            (when (not tli) (vbasense--info "can't resolve GUID[%s] from ProgID[%s]" guid progid))
            (when (stringp vartype)
              (setq vbasense--current-instance-alist
                    (delete currdef vbasense--current-instance-alist))
              (push `(,varnm . ,vartype) vbasense--current-instance-alist)
              (setq srchvars (delete varnm srchvars))
              (vbasense--trace "embodied instance[%s] to [%s] by createobject statement" varnm vartype))))))
    ;; find substitute statement
    (when srchvars
      (goto-char (point-min))
      (let ((re (rx-to-string `(and bol (* space)
                                    (or ,@vbasense--words-set) (+ space)
                                    (group (or ,@srchvars)) (+ space) "=" (+ space)
                                    (group (+ not-newline)) eol))))
        (while (re-search-forward re nil t)
          (let* ((varnm (match-string-no-properties 1))
                 (subst (match-string-no-properties 2))
                 (currdef (assoc varnm vbasense--current-instance-alist)))
            (when (not (string-match "\\`[cC]reate[oO]bject(" subst))
              (skip-syntax-backward "-")
              (multiple-value-bind (nstype ns) (vbasense--get-namespace-of-current-context)
                (when (and (or (eq nstype 'class)
                               (eq nstype 'interface))
                           (stringp ns))
                  (setq vbasense--current-instance-alist
                        (delete currdef vbasense--current-instance-alist))
                  (push `(,varnm . ,ns) vbasense--current-instance-alist)
                  (setq srchvars (delete varnm srchvars))
                  (vbasense--trace "embodied instance[%s] to [%s] by substitute statement" varnm ns))))))))
    ;; find for statement
    (when srchvars
      (goto-char startpt)
      (let ((re (rx-to-string `(and bol (* space)
                                    (or ,@vbasense--words-for) (+ space)
                                    (or ,@vbasense--words-each) (+ space)
                                    (group (or ,@srchvars)) (+ space)
                                    (or ,@vbasense--words-in) (+ space)
                                    (+ not-newline) eol))))
        (while (re-search-forward re endpt t)
          (let* ((varnm (match-string-no-properties 1))
                 (currdef (assoc varnm vbasense--current-instance-alist)))
            (skip-syntax-backward "-")
            (multiple-value-bind (nstype ns) (vbasense--get-namespace-of-current-context)
              (when (and (or (eq nstype 'class)
                             (eq nstype 'interface))
                         (stringp ns))
                (let* ((mtd (vbasense--get-latest-method-definition ns nstype "Item"))
                       (ret (when (vbasense--method-p mtd)
                              (vbasense--method-ret mtd))))
                  (when (stringp ret)
                    (setq vbasense--current-instance-alist
                          (delete currdef vbasense--current-instance-alist))
                    (push `(,varnm . ,ret) vbasense--current-instance-alist)
                    (setq srchvars (delete varnm srchvars))
                    (vbasense--trace "embodied instance[%s] to [%s] by for statement" varnm ret)))))))))
    ))


;;;;;;;;;;;;;;;;;;;;;
;; Reading Context

(defsubst vbasense--get-available-classes-on-current ()
  (append vbasense--available-classes
          vbasense--current-classes))

(defsubst vbasense--get-available-enums-on-current ()
  (append vbasense--available-enums
          vbasense--current-enums))

(defun vbasense--jump-current-state-head ()
  (beginning-of-line)
  (when (loop while (not (bobp))
              do (forward-line -1)
              if (not (string-match "_\\s-*$" (thing-at-point 'line))) return t)
    (forward-line 1))
  (skip-syntax-forward "-"))

(defun vbasense--get-implicit-namespace-method-definition (methodnm &optional methodtype)
  (loop for k being the hash-keys in vbasense--hash-object-cache using (hash-values tli)
        for mtd = (when (and (string-match "\\`[mc]:.+\\..+\\'" k)
                             (or (vbasense--module-p tli)
                                 (and (vbasense--class-p tli)
                                      (vbasense--class-coclassp tli))))
                        (vbasense--get-method-of-tli tli methodnm methodtype))
        if mtd return mtd))

;; (defsubst vbasense--get-current-valid-line (&optional pos)
;;   (save-excursion
;;     (let* ((ret (buffer-substring-no-properties (or pos
;;                                                     (progn (end-of-line) (point)))
;;                                                 (progn (beginning-of-line) (point)))))
;;       (setq ret (replace-regexp-in-string "^\\s-+" "" ret))
;;       (setq ret (replace-regexp-in-string "\\s-+$" "" ret))
;;       (setq ret (replace-regexp-in-string "_$" "" ret))
;;       (setq ret (replace-regexp-in-string "\n" "" ret))
;;       ret)))

;; (defun vbasense--get-current-statement ()
;;   (save-excursion
;;     (let* ((ret (vbasense--get-current-valid-line (point))))
;;       (ignore-errors
;;         (while (and (not (bobp))
;;                     (forward-line -1)
;;                     (string-match "_\\s-*$" (thing-at-point 'line)))
;;           (setq ret (concat (vbasense--get-current-valid-line) ret))))
;;       ret)))

(defun vbasense--get-current-context ()
  (yaxception:$
    (yaxception:try
      (save-excursion
        (let ((word (buffer-substring-no-properties (point)
                                                    (progn (skip-syntax-backward "w_") (point))))
              (conn-char (or (ignore-errors (forward-char -1)
                                            (format "%c" (char-after)))
                             "")))
          (cond ((string= conn-char " ")
                 (let ((pre-word (downcase
                                  (save-excursion (buffer-substring-no-properties
                                                   (progn (skip-syntax-backward "-") (point))
                                                   (progn (skip-syntax-backward "w_") (point)))))))
                   (cond ((string= word "_")
                          (list 'lineend))
                         ((string-match vbasense--regexp-word-of-class pre-word)
                          (cond ((string= pre-word "as") (list 'class '("New")))
                                (t                       (list 'class))))
                         ((string-match vbasense--regexp-word-of-implement pre-word)
                          (list 'implement))
                         ((string-match vbasense--regexp-word-of-variable pre-word)
                          (cond ((string= pre-word "redim") (list 'variable '("Preserve")))
                                (t                          (list 'variable))))
                         ((string-match vbasense--regexp-word-of-scope pre-word)
                          (list 'modifier '("Dim" "Static" "Const" "Function" "Sub" "Get" "Let" "Type" "Enum" "ByVal" "ByRef" "Declare")))
                         ((string-match vbasense--regexp-word-of-exit pre-word)
                          (list 'modifier '("Sub" "Function" "For" "Do")))
                         ((string-match vbasense--regexp-word-of-end pre-word)
                          (list 'modifier '("Sub" "Function" "Property" "Type" "Enum" "If" "Select" "With")))
                         ((string-match vbasense--regexp-word-of-optional pre-word)
                          (list 'modifier '("ByVal" "ByRef")))
                         ((string-match vbasense--regexp-word-of-property pre-word)
                          (list 'modifier '("Get" "Let")))
                         ((string-match vbasense--regexp-word-of-procedure-modifier pre-word)
                          (list 'modifier '("Sub" "Function" "Property")))
                         ((string-match vbasense--regexp-word-of-definition pre-word)
                          (list 'definition))
                         ((string-match vbasense--regexp-word-of-label pre-word)
                          (list 'label))
                         (t
                          (ignore-errors (forward-char 1)) ;move to param context
                          (list 'param (vbasense--get-param-type-of-current-context))))))
                ((string= conn-char ".")
                 (multiple-value-bind (nstype ns) (vbasense--get-namespace-of-current-context)
                   (case nstype
                     (app       (list 'app-member ns))
                     (module    (list 'module-member ns))
                     (interface (list 'interface-member ns))
                     (class     (list 'class-member ns))
                     (enum      (list 'enum-member ns))
                     (t         (list 'unknown-member)))))
                ((or (string= conn-char "(")
                     (string= conn-char ","))
                 (ignore-errors (forward-char 1)) ;move to param context
                 (list 'param (vbasense--get-param-type-of-current-context)))))))
    (yaxception:catch 'error e
      (vbasense--error "failed get current context : %s\n%s"
                       (yaxception:get-text e) (yaxception:get-stack-trace-string e))
      nil)))

(defun vbasense--get-current-identifier ()
  (yaxception:$
    (yaxception:try
      (save-excursion
        (let* ((word (buffer-substring-no-properties
                      (progn (skip-syntax-forward "w_") (point))
                      (progn (skip-syntax-backward "w_") (point))))
               (conn-char (or (when (not (string= word ""))
                                (ignore-errors (forward-char -1)
                                               (format "%c" (char-after))))
                              "")))
          (cond
           ;; current word is member of something
           ((string= conn-char ".")
            (multiple-value-bind (nstype ns) (vbasense--get-namespace-of-current-context)
              (case nstype
                (app (let* ((fullnsl (downcase (concat ns "." word))))
                       (or (gethash (concat "m:" fullnsl) vbasense--hash-object-cache)
                           (gethash (concat "i:" fullnsl) vbasense--hash-object-cache)
                           (gethash (concat "c:" fullnsl) vbasense--hash-object-cache)
                           (gethash fullnsl vbasense--hash-enum-cache))))
                (t (or (vbasense--get-latest-method-definition ns nstype word)
                       (when (eq nstype 'enum)
                         (let ((enum (vbasense--get-latest-enum-definition ns)))
                           (when enum
                             (loop for m in (vbasense--enum-members enum)
                                   if (string= (vbasense--enummember-name m) word) return m))))
                       (when (eq nstype 'module)
                         (let* ((fullnsl (downcase (concat ns "." word))))
                           (or (gethash (concat "c:" fullnsl) vbasense--hash-object-cache)
                               (gethash fullnsl vbasense--hash-enum-cache)))))))))
           ;; current word is instance of current
           ((member word vbasense--current-instances)
            (let* ((maybe (assoc-default word vbasense--current-instance-alist))
                   (maybel (when (stringp maybe) (downcase maybe))))
              (cond ((member maybe vbasense--current-classes)
                     (vbasense--get-current-definition maybe 'class))
                    ((member maybe vbasense--current-enums)
                     (vbasense--get-current-definition maybe 'enum))
                    (t
                     (or (gethash (concat "m:" maybel) vbasense--hash-object-cache)
                         (gethash (concat "i:" maybel) vbasense--hash-object-cache)
                         (gethash (concat "c:" maybel) vbasense--hash-object-cache)
                         (gethash maybel vbasense--hash-enum-cache))))))
           ;; current word is method of current
           ((member word vbasense--current-methods)
            (vbasense--get-current-definition word 'method))
           ;; current word is bare-word of current class
           ((member word vbasense--current-classes)
            (vbasense--get-current-definition word 'class))
           ;; current word is bare-word of current enum
           ((member word vbasense--current-enums)
            (vbasense--get-current-definition word 'enum))
           ;; current word is method of implicit namespace
           ((member word vbasense--available-implicit-methods)
            (vbasense--get-implicit-namespace-method-definition word))
           ;; current word is bare-word of application
           ((member word vbasense--available-applications)
            (gethash (downcase word) vbasense--hash-app-cache))
           ;; current word is bare-word of module
           ((member word vbasense--available-modules)
            (gethash (concat "m:" (downcase word)) vbasense--hash-object-cache))
           ;; current word is bare-word of class
           ((member word vbasense--available-classes)
            (gethash (concat "c:" (downcase word)) vbasense--hash-object-cache))
           ;; current word is bare-word of enum
           ((member word vbasense--available-enums)
            (gethash (downcase word) vbasense--hash-enum-cache))))))
    (yaxception:catch 'error e
      (vbasense--error "failed get current identifier : %s\n%s"
                       (yaxception:get-text e) (yaxception:get-stack-trace-string e))
      nil)))

(defsubst vbasense--jump-with-state-tail ()
  (vbasense--trace "start jump with state tail")
  (when (re-search-backward "^\\s-*[wW]ith +" nil t)
    (goto-char (point-at-eol))
    (skip-syntax-backward "-")
    t))

(defun vbasense--get-namespace-of-current-context ()
  (yaxception:$
    (yaxception:try
      (save-excursion
        (let* ((last-char (format "%c" (char-before)))
               (word (buffer-substring-no-properties
                      (progn (when (string= last-char ")") (backward-sexp)) (point))
                      (progn (skip-syntax-backward "w_") (point))))
               (pre-char (or (ignore-errors (forward-char -1)
                                            (format "%c" (char-after)))
                             ""))
               (maybe (assoc-default word vbasense--current-instance-alist))
               (maybel (when (stringp maybe) (downcase maybe))))
          (cond 
           ;; current word is member of something
           ((and (not (string= word ""))
                 (string= pre-char "."))
            (multiple-value-bind (nstype ns) (vbasense--get-namespace-of-current-context)
              (case nstype
                (app (let* ((fullns (concat ns "." word))
                            (fullnsl (downcase fullns)))
                       ;; current word is member of application
                       (cond ((gethash (concat "m:" fullnsl) vbasense--hash-object-cache) (list 'module fullns))
                             ((gethash (concat "i:" fullnsl) vbasense--hash-object-cache) (list 'interface fullns))
                             ((gethash (concat "c:" fullnsl) vbasense--hash-object-cache) (list 'class fullns))
                             ((gethash fullnsl vbasense--hash-enum-cache)                 (list 'enum fullns))
                             (t
                              (vbasense--info "can't find member[%s] of application[%s]" word ns)))))
                (t (let* ((mtd (vbasense--get-latest-method-definition ns nstype word))
                          (ret (when (vbasense--method-p mtd)
                                 (vbasense--method-ret mtd))))
                     (cond ((stringp ret)
                            ;; current word is method of explicit namespace
                            (let ((retl (downcase ret)))
                              (cond ((gethash (concat "i:" retl) vbasense--hash-object-cache) (list 'interface ret))
                                    ((gethash (concat "c:" retl) vbasense--hash-object-cache) (list 'class ret))
                                    ((gethash retl vbasense--hash-enum-cache)                 (list 'enum ret))
                                    ((member ret vbasense--current-classes)                   (list 'class ret))
                                    ((member ret vbasense--current-enums)                     (list 'enum ret))
                                    (t
                                     (vbasense--info "can't find return[%s] of nstype[%s] ns[%s] method[%s]" nstype ns word)))))
                           ((eq nstype 'module)
                            ;; current word is member of module
                            (let* ((fullns (concat ns "." word))
                                   (fullnsl (downcase fullns)))
                              (cond ((gethash (concat "c:" fullnsl) vbasense--hash-object-cache) (list 'class fullns))
                                    ((gethash fullnsl vbasense--hash-enum-cache)                 (list 'enum fullns))
                                    (t
                                     (vbasense--info "can't find member[%s] of module[%s]" word ns)))))))))))
           ;; current word is instance
           ((stringp maybe)
            (cond ((gethash (concat "m:" maybel) vbasense--hash-object-cache) (list 'module maybe))
                  ((gethash (concat "i:" maybel) vbasense--hash-object-cache) (list 'interface maybe))
                  ((gethash (concat "c:" maybel) vbasense--hash-object-cache) (list 'class maybe))
                  ((gethash maybel vbasense--hash-enum-cache)                 (list 'enum maybe))
                  ((member maybe vbasense--current-classes)                   (list 'class maybe))
                  ((member maybe vbasense--current-enums)                     (list 'enum maybe))
                  (t
                   (vbasense--info "can't find type[%s] of instance[%s]" maybe word))))
           ;; current is top level in with-statement
           ((and (string= word "")
                 (or (string-match "\\`\\s-\\'" pre-char)
                     (string= pre-char "("))
                 (vbasense--jump-with-state-tail))
            (vbasense--get-namespace-of-current-context))
           ;; current word is bare-word of application
           ((member word vbasense--available-applications) (list 'app word))
           ;; current word is bare-word of module
           ((member word vbasense--available-modules) (list 'module word))
           ;; current word is bare-word of coclass
           ((member word vbasense--available-coclasses) (list 'class word))
           ;; current word is bare-word of enum
           ((member word (vbasense--get-available-enums-on-current)) (list 'enum word))
           ;; current word is method of current
           ((member word vbasense--current-methods)
            (let* ((method (vbasense--get-current-definition word 'function))
                   (ret (when method (vbasense--method-ret method))))
              (cond ((member ret vbasense--available-interfaces)               (list 'interface ret))
                    ((member ret (vbasense--get-available-classes-on-current)) (list 'class ret))
                    ((member ret (vbasense--get-available-enums-on-current))   (list 'enum ret))
                    (t
                     (vbasense--info "can't find return[%s] of current method[%s]" ret word)))))
           ;; current word is method of implicit namespace
           ((member word vbasense--available-implicit-methods)
            (let* ((mtd (vbasense--get-implicit-namespace-method-definition word))
                   (ret (when (vbasense--method-p mtd)
                          (vbasense--method-ret mtd))))
              (cond ((member ret vbasense--available-interfaces)               (list 'interface ret))
                    ((member ret (vbasense--get-available-classes-on-current)) (list 'class ret))
                    ((member ret (vbasense--get-available-enums-on-current))   (list 'enum ret))
                    (t
                     (vbasense--info "can't find return[%s] of implicit method[%s]" ret word)))))
           (t
            (vbasense--info "can't find namespace : last-char[%s] word[%s] pre-char[%s]" last-char word pre-char))))))
    (yaxception:catch 'error e
      (vbasense--error "failed get namespace of current context : %s\n%s"
                       (yaxception:get-text e) (yaxception:get-stack-trace-string e))
      nil)))
    
(defun vbasense--get-current-param-startpt ()
  (save-excursion
    (let* ((currpt (point))
           (toppt (save-excursion (vbasense--jump-current-state-head) (point)))
           (paren-startpt (loop with depth = 1
                                for openpt = (or (save-excursion
                                                   (ignore-errors
                                                     (when (search-backward "(" toppt t) (point))))
                                                 toppt)
                                for closept = (or (save-excursion
                                                    (ignore-errors
                                                      (when (search-backward ")" toppt t) (point))))
                                                  toppt)
                                for nextpt = (cond ((> openpt closept)
                                                    (decf depth)
                                                    openpt)
                                                   (t
                                                    (incf depth)
                                                    closept))
                                while (> (point) toppt)
                                do (goto-char nextpt)
                                if (= depth 0) return (point))))
      (or paren-startpt
          (when (search-forward " " currpt t)
            (forward-char -1)
            (point))))))

(defun vbasense--get-current-param-index (param-text)
  (if (not (stringp param-text))
      -1
    (setq param-text (replace-regexp-in-string "([^(]*)" "" param-text))
    (- (length (split-string param-text ",")) 1)))

(defun vbasense--get-param-type-of-current-context ()
  (yaxception:$
    (yaxception:try
      (save-excursion
        (let* ((paramstartpt (vbasense--get-current-param-startpt))
               (paramidx (when paramstartpt
                           (vbasense--get-current-param-index
                            (buffer-substring-no-properties (+ paramstartpt 1) (point)))))
               (methodnm (when paramstartpt
                           (goto-char paramstartpt)
                           (buffer-substring-no-properties (point)
                                                           (progn (skip-syntax-backward "w_") (point)))))
               (pre-char (or (ignore-errors (forward-char -1)
                                            (format "%c" (char-after)))
                             ""))
               (method (cond 
                        ;; current method is belonged class
                        ((string= pre-char ".")
                         (multiple-value-bind (nstype ns) (vbasense--get-namespace-of-current-context)
                           (vbasense--get-latest-method-definition ns nstype methodnm)))
                        ;; current method of implicit namespace
                        ((member methodnm vbasense--available-implicit-methods)
                         (vbasense--get-implicit-namespace-method-definition methodnm))
                        ;; current word is method of current
                        ((member methodnm vbasense--current-methods)
                         (vbasense--get-current-definition methodnm 'method))))
               (param (when (vbasense--method-p method)
                        (nth paramidx (vbasense--method-params method)))))
          (or (when (vbasense--param-p param) (vbasense--param-type param))
              (progn (vbasense--info "can't find param : paramstartpt[%s] paramidx[%s] methodnm[%s] pre-char[%s]"
                                     paramstartpt paramidx methodnm pre-char)
                     "")))))
    (yaxception:catch 'error e
      (vbasense--error "failed get param type of current context : %s\n%s"
                       (yaxception:get-text e) (yaxception:get-stack-trace-string e))
      "")))


;;;;;;;;;;;;;;;;;;;;;;;
;; For auto-complete

(defun vbasense--insert-with-ac-trigger-command (n)
  (interactive "p")
  (self-insert-command n)
  (auto-complete-1 :triggered 'trigger-key))

(defvar vbasense--ac-candidate-type-alist nil)
(defsubst vbasense--set-ac-cand-type (cand candtype)
  (let ((candl (downcase cand)))
    (when (not (assoc-default candl vbasense--ac-candidate-type-alist))
      (push `(,candl . ,candtype) vbasense--ac-candidate-type-alist)))
  cand)

(defsubst vbasense--set-ac-cands-type (cands candtype)
  (loop for c in cands
        collect (vbasense--set-ac-cand-type c candtype)))

(defsubst vbasense--get-methods-in-namespace (ns nstype)
  (let* ((tli (case nstype
                (module    (gethash (concat "m:" (downcase ns)) vbasense--hash-object-cache))
                (interface (gethash (concat "i:" (downcase ns)) vbasense--hash-object-cache))
                (class     (vbasense--get-latest-class-definition ns))))
         (methods (cond ((vbasense--module-p tli)    (vbasense--module-methods tli))
                        ((vbasense--interface-p tli) (vbasense--interface-methods tli))
                        ((vbasense--class-p tli)     (vbasense--class-methods tli)))))
    (loop for method in methods
          collect (vbasense--set-ac-cand-type (vbasense--method-name method) 'nsmethod))))

(defsubst vbasense--get-app-members (appnm)
  (let ((app (gethash (downcase appnm) vbasense--hash-app-cache)))
    (when app
      (append (loop for m in (vbasense--app-modules app)
                    collect (vbasense--set-ac-cand-type (vbasense--module-name m) 'module))
              (loop for i in (vbasense--app-interfaces app)
                    collect (vbasense--set-ac-cand-type (vbasense--interface-name i) 'interface))
              (loop for c in (vbasense--app-classes app)
                    collect (vbasense--set-ac-cand-type (vbasense--class-name c) 'class))
              (loop for e in (vbasense--app-enums app)
                    collect (vbasense--set-ac-cand-type (vbasense--enum-name e) 'enum))))))

(defsubst vbasense--get-module-members (modnm)
  (let ((mod (gethash (concat "m:" (downcase modnm)) vbasense--hash-object-cache)))
    (when mod
      (append (loop for c in (vbasense--module-classes mod)
                    collect (vbasense--set-ac-cand-type (vbasense--class-name c) 'class))
              (loop for e in (vbasense--module-enums mod)
                    collect (vbasense--set-ac-cand-type (vbasense--enum-name e) 'enum))
              (vbasense--get-methods-in-namespace modnm 'module)))))

(defsubst vbasense--get-enum-members (enumnm)
  (let ((enum (vbasense--get-latest-enum-definition enumnm)))
    (when enum
      (loop for m in (vbasense--enum-members enum)
            collect (vbasense--set-ac-cand-type (vbasense--enummember-name m) 'enummember)))))

(defsubst vbasense--get-instances-with-type (type)
  (loop for i in vbasense--current-instances
        for currtype = (assoc-default i vbasense--current-instance-alist)
        if (and (stringp currtype)
                (string= currtype type))
        collect (vbasense--set-ac-cand-type i 'var)))

(defsubst vbasense--get-instancable-idents ()
  (append (vbasense--set-ac-cands-type vbasense--builtin-objects 'bobject)
          (vbasense--set-ac-cands-type vbasense--available-applications 'app)
          (vbasense--set-ac-cands-type vbasense--available-modules 'module)
          (vbasense--set-ac-cands-type vbasense--available-interfaces 'interface)
          (vbasense--set-ac-cands-type vbasense--current-classes 'currclass)
          (vbasense--set-ac-cands-type vbasense--available-classes 'class)
          (vbasense--set-ac-cands-type vbasense--current-enums 'currenum)
          (vbasense--set-ac-cands-type vbasense--available-enums 'enum)))

(defsubst vbasense--get-bareable-idents ()
  (append (vbasense--set-ac-cands-type vbasense--builtin-keywords 'keyword)
          (vbasense--set-ac-cands-type vbasense--current-instances 'var)
          (vbasense--set-ac-cands-type vbasense--available-applications 'app)
          (vbasense--set-ac-cands-type vbasense--available-modules 'module)
          (vbasense--set-ac-cands-type vbasense--available-coclasses 'class)
          (vbasense--set-ac-cands-type vbasense--current-enums 'currenum)
          (vbasense--set-ac-cands-type vbasense--available-enums 'enum)
          (vbasense--set-ac-cands-type vbasense--builtin-functions 'bmethod)
          (vbasense--set-ac-cands-type vbasense--current-methods 'currmethod)
          (vbasense--set-ac-cands-type vbasense--available-implicit-methods 'nsmethod)))

(defvar vbasense--ac-action-function nil)
(defun vbasense--ac-action-function ()
  (when (functionp vbasense--ac-action-function)
    (funcall vbasense--ac-action-function)))

(defun vbasense--ac-action-for-implements ()
  (when (y-or-n-p "[VBASense] Insert definition of the implement interface procedure?")
    (insert "\n\n")
    (vbasense-insert-implement-definition)))

(defun vbasense--get-ac-candidates ()
  (yaxception:$
    (yaxception:try
      (vbasense--trace "start get ac candidates")
      (if (or (= vbasense-lookup-current-buffer-threshold 0)
              (< (point-max) vbasense-lookup-current-buffer-threshold))
          (vbasense--update-current-definition)
        (vbasense--trace "skip update current definition : buffersize[%s]" (point-max))
        (vbasense--update-current-instance))
      (setq vbasense--ac-candidate-type-alist nil)
      (setq vbasense--ac-action-function nil)
      (multiple-value-bind (ctx ident-value) (vbasense--get-current-context)
        (vbasense--debug "got current context. ctx[%s] ident-value[%s]" ctx ident-value)
        (case ctx
          (lineend          nil)
          (definition       nil)
          (label            nil)
          (modifier         (vbasense--set-ac-cands-type ident-value 'keyword))
          (class            (append (vbasense--set-ac-cands-type ident-value 'keyword)
                                    (vbasense--get-instancable-idents)))
          (implement        (setq vbasense--ac-action-function 'vbasense--ac-action-for-implements)
                            (vbasense--set-ac-cands-type vbasense--available-interfaces 'interface))
          (variable         (append (vbasense--set-ac-cands-type ident-value 'keyword)
                                    (vbasense--set-ac-cands-type vbasense--current-instances 'var)))
          (app-member       (vbasense--get-app-members ident-value))
          (module-member    (vbasense--get-module-members ident-value))
          (interface-member (vbasense--get-methods-in-namespace ident-value 'interface))
          (class-member     (vbasense--get-methods-in-namespace ident-value 'class))
          (enum-member      (vbasense--get-enum-members ident-value))
          (unknown-member   nil)
          (param            (cond ((member ident-value (vbasense--get-available-enums-on-current))
                                   (append (vbasense--get-instances-with-type ident-value)
                                           (vbasense--get-enum-members ident-value)
                                           (list (vbasense--set-ac-cand-type ident-value 'enum))))
                                  (t
                                   (vbasense--get-bareable-idents))))
          (t                (append (vbasense--get-bareable-idents)
                                    (vbasense--set-ac-cands-type vbasense--builtin-objects 'bobject))))))
    (yaxception:catch 'error e
      (vbasense--error "failed get ac candidates : %s\n%s"
                       (yaxception:get-text e) (yaxception:get-stack-trace-string e))
      nil)))

(defun vbasense--get-ac-progid-candidates ()
  (vbasense--set-ac-cands-type vbasense--available-progids 'progid))

(defun vbasense--get-ac-implicit-enummember-candidates ()
  (vbasense--set-ac-cands-type vbasense--available-implicit-enummembers 'enummember))

(defun* vbasense--get-any-document (&key x xname xtype)
  (let* ((xtype (or xtype
                    (cond ((vbasense--app-p x)        'app)
                          ((vbasense--module-p x)     'module)
                          ((vbasense--interface-p x)  'interface)
                          ((vbasense--class-p x)      'class)
                          ((vbasense--enum-p x)       'enum)
                          ((vbasense--enummember-p x) 'enummember)
                          ((vbasense--method-p x)     'method))))
         (typenm (case xtype
                   (app                          "Application")
                   (module                       "Module")
                   (interface                    "Interface")
                   ((class currclass)            "Class")
                   ((enum currenum)              "Enum")
                   (enummember                   "Member of Enum")
                   (var                          "Variable")
                   ((method nsmethod currmethod) "Procedure")
                   (bmethod                      "Builtin Function")
                   (bobject                      "Builtin Object")
                   (progid                       "Progammable Identifier")
                   (keyword                      "Keyword of Visual Basic")
                   (t                            "Unknown")))
         (xname (or xname
                    (cond ((vbasense--app-p x)        (vbasense--app-name x))
                          ((vbasense--module-p x)     (vbasense--module-name x))
                          ((vbasense--interface-p x)  (vbasense--interface-name x))
                          ((vbasense--class-p x)      (vbasense--class-name x))
                          ((vbasense--enum-p x)       (vbasense--enum-name x))
                          ((vbasense--enummember-p x) (vbasense--enummember-name x))
                          ((vbasense--method-p x)     (vbasense--method-name x)))))
         (helptext (cond ((vbasense--module-p x)     (vbasense--module-helptext x))
                         ((vbasense--interface-p x)  (vbasense--interface-helptext x))
                         ((vbasense--class-p x)      (vbasense--class-helptext x))
                         ((vbasense--enum-p x)       (vbasense--enum-helptext x))
                         ((vbasense--enummember-p x) (vbasense--enummember-helptext x))
                         ((vbasense--method-p x)     (vbasense--method-helptext x))))
         (helpfile (cond ((vbasense--module-p x)     (vbasense--module-helpfile x))
                         ((vbasense--interface-p x)  (vbasense--interface-helpfile x))
                         ((vbasense--class-p x)      (vbasense--class-helpfile x))
                         ((vbasense--enum-p x)       (vbasense--enum-helpfile x))
                         ((vbasense--enummember-p x) (vbasense--enummember-helpfile x))
                         ((vbasense--method-p x)     (vbasense--method-helpfile x))))
         (helpfile (cond ((or (not (stringp helpfile))
                              (string= helpfile ""))   "Nothing")
                         (t                            (concat "'" helpfile "'")))))
    (when (or (not (stringp helptext))
              (string= helptext ""))
      (setq helptext "Not documented."))
    (when (vbasense--method-p x)
      (setq helptext (concat helptext "\n"
                             "\n"
                             "[Signature]\n"
                             (vbasense--get-method-eldoc x))))
    (concat xname " is " typenm ".\n"
            "\n"
            helptext "\n"
            "\n"
            "HelpFile is " helpfile ".\n")))

(defun vbasense--get-ac-document (cand)
  (yaxception:$
    (yaxception:try
      (set-text-properties 0 (string-width cand) nil cand)
      (let* ((candl (downcase cand))
             (ctype (assoc-default candl vbasense--ac-candidate-type-alist))
             (cobj (case ctype
                     (module    (gethash (concat "m:" candl) vbasense--hash-object-cache))
                     (interface (gethash (concat "i:" candl) vbasense--hash-object-cache))
                     (class     (gethash (concat "c:" candl) vbasense--hash-object-cache))
                     (enum      (gethash candl vbasense--hash-enum-cache))
                     (progid    (let ((guid (gethash candl vbasense--hash-tli-prog-cache)))
                                  (gethash guid vbasense--hash-tli-guid-cache))))))
        (vbasense--get-any-document :xtype ctype :xname cand :x cobj)))
    (yaxception:catch 'error e
      (vbasense--error "failed get ac document : %s\n%s"
                       (yaxception:get-text e) (yaxception:get-stack-trace-string e))
      "")))

(defvar ac-source-vbasense-head
  '((candidates . vbasense--get-ac-candidates)
    (prefix . "^\\s-*\\([a-zA-Z][a-zA-Z_0-9]*\\)")
    (symbol . "h")
    (document . vbasense--get-ac-document)
    (requires . 1)
    (action . vbasense--ac-action-function)
    (cache)
    (limit . 500)))

(defvar ac-source-vbasense-word
  '((candidates . vbasense--get-ac-candidates)
    (prefix . "[a-zA-Z][a-zA-Z_0-9]*\\s-+\\([a-zA-Z_0-9]*\\)")
    (symbol . "w")
    (document . vbasense--get-ac-document)
    (requires . 0)
    (action . vbasense--ac-action-function)
    (cache)
    (limit . 500)))

(defvar ac-source-vbasense-operator
  '((candidates . vbasense--get-ac-candidates)
    (prefix . "\\s-+\\(?:=\\|=?<=?\\|[=<]?>=?\\|&\\)\\s-+\\([a-zA-Z_0-9]*\\)")
    (symbol . "o")
    (document . vbasense--get-ac-document)
    (requires . 0)
    (action . vbasense--ac-action-function)
    (cache)
    (limit . 500)))

(defvar ac-source-vbasense-dot
  '((candidates . vbasense--get-ac-candidates)
    (prefix . "\\.\\([a-zA-Z_0-9]*\\)")
    (symbol . "m")
    (document . vbasense--get-ac-document)
    (requires . 0)
    (action . vbasense--ac-action-function)
    (cache)))

(defvar ac-source-vbasense-paren
  '((candidates . vbasense--get-ac-candidates)
    (prefix . "(\\s-*\\([a-zA-Z_0-9]*\\)")
    (symbol . "p")
    (document . vbasense--get-ac-document)
    (requires . 0)
    (action . vbasense--ac-action-function)
    (cache)
    (limit . 500)))

(defvar ac-source-vbasense-comma
  '((candidates . vbasense--get-ac-candidates)
    (prefix . ",\\s-+\\([a-zA-Z_0-9]*\\)")
    (symbol . "c")
    (document . vbasense--get-ac-document)
    (requires . 0)
    (action . vbasense--ac-action-function)
    (cache)
    (limit . 500)))

(defvar ac-source-vbasense-keywordargument
  '((candidates . vbasense--get-ac-candidates)
    (prefix . "[a-zA-Z][a-zA-Z_0-9]*:=\\([a-zA-Z_0-9]*\\)")
    (symbol . "k")
    (document . vbasense--get-ac-document)
    (requires . 0)
    (action . vbasense--ac-action-function)
    (cache)
    (limit . 500)))

(defvar ac-source-vbasense-createobject
  '((candidates . vbasense--get-ac-progid-candidates)
    (prefix . "\\<[cC]reate[oO]bject(\"?\\([a-zA-Z0-9_.]*\\)")
    (symbol . "x")
    (document . vbasense--get-ac-document)
    (requires . 0)
    (action . (lambda ()
                (let ((currpt (point)))
                  (beginning-of-line)
                  (or (when (re-search-forward "\\<[cC]reate[oO]bject(\"?\\([a-zA-Z0-9_.]+\\)\"?)?" nil t)
                        (goto-char (match-end 0))
                        (replace-match (concat "CreateObject(\"" (match-string-no-properties 1) "\")")))
                      (goto-char currpt)))))
    (cache)
    (limit . 500)))

(defvar ac-source-vbasense-implicit-enummember nil)
(defun vbasense--get-ac-source-implicit-enummember ()
  (let ((re (rx-to-string `(and (not (any ".")) bow
                                (group (or ,@vbasense-ac-implicit-enummember-prefixes)
                                       (* (any "a-zA-Z_0-9")))))))
    `((candidates . vbasense--get-ac-implicit-enummember-candidates)
      (prefix . ,re)
      (symbol . "i")
      (document . vbasense--get-ac-document)
      (requires . 0)
      (cache)
      (limit . 500))))


;;;;;;;;;;;;;;;
;; For eldoc

(defun vbasense--echo-method-usage ()
  (yaxception:$
    (yaxception:try
      (vbasense--update-current-instance)
      (princ (vbasense--get-current-method-eldoc)))
    (yaxception:catch 'error e
      (vbasense--error "failed echo method usage : %s\n%s"
                       (yaxception:get-text e) (yaxception:get-stack-trace-string e))
      (princ ""))))

(defvar vbasense--last-method-eldoc "")
(make-variable-buffer-local 'vbasense--last-method-eldoc)
(defvar vbasense--last-param-startpt 0)
(make-variable-buffer-local 'vbasense--last-param-startpt)
(defvar vbasense--last-param-index -2)
(make-variable-buffer-local 'vbasense--last-param-index)
(defvar vbasense--last-method-definition nil)
(make-variable-buffer-local 'vbasense--last-method-definition)
(defun vbasense--get-current-method-eldoc ()
  (let* ((currface (get-text-property (point) 'face))
         (paramstartpt (when (not (eq currface 'font-lock-comment-face))
                         (vbasense--get-current-param-startpt)))
         (paramidx (when paramstartpt
                     (vbasense--get-current-param-index
                      (buffer-substring-no-properties (+ paramstartpt 1) (point))))))
    (cond ((not paramstartpt)
           "")
          ((and (eq paramstartpt vbasense--last-param-startpt)
                (eq paramidx vbasense--last-param-index))
           vbasense--last-method-eldoc)
          ((and (eq paramstartpt vbasense--last-param-startpt)
                vbasense--last-method-definition)
           (setq vbasense--last-param-index paramidx)
           (setq vbasense--last-method-eldoc
                 (vbasense--get-method-eldoc vbasense--last-method-definition paramidx)))
          (t
           (vbasense--trace "start get current method eldoc : startpt[%s] idx[%s]" paramstartpt paramidx)
           (save-excursion
             (goto-char paramstartpt)
             (ignore-errors (forward-char -1))
             (setq vbasense--last-param-startpt paramstartpt)
             (setq vbasense--last-param-index paramidx)
             (setq vbasense--last-method-definition nil)
             (setq vbasense--last-method-eldoc
                   (let* ((x (vbasense--get-current-identifier))
                          (xname (cond ((vbasense--app-p x)        (vbasense--app-name x))
                                       ((vbasense--module-p x)     (vbasense--module-name x))
                                       ((vbasense--interface-p x)  (vbasense--interface-name x))
                                       ((vbasense--class-p x)      (vbasense--class-name x))
                                       ((vbasense--enum-p x)       (vbasense--enum-name x))
                                       ((vbasense--enummember-p x) (vbasense--enummember-name x))
                                       ((vbasense--method-p x)     (vbasense--method-name x)))))
                     (cond ((vbasense--method-p x)
                            (setq vbasense--last-method-definition x)
                            (vbasense--get-method-eldoc x paramidx))
                           (t
                            (vbasense--info "can't identify method from %s" xname)
                            "")))))))))

(defsubst vbasense--get-method-param-strings (mtd &optional unknown-type)
  (loop for p in (vbasense--method-params mtd)
        collect (concat (cond ((vbasense--param-optional p) "Optional ")
                              ((vbasense--param-arrayp p)   "ParamArray ")
                              (t                            ""))
                        (if (vbasense--param-ref p) "ByRef " "ByVal ")
                        (vbasense--param-name p)
                        " As "
                        (cond ((and (stringp (vbasense--param-type p))
                                    (not (string= (vbasense--param-type p) ""))) (vbasense--param-type p))
                              (unknown-type                                      unknown-type)
                              (t                                                 "=Unknown="))
                        (if (string= (vbasense--param-default p) "")
                            ""
                          (concat " = " (vbasense--param-default p))))))

(defsubst vbasense--get-method-ret-string (mtd)
  (if (and (stringp (vbasense--method-ret mtd))
           (not (string= (vbasense--method-ret mtd) "")))
      (concat " As " (vbasense--method-ret mtd))
    ""))

(defun vbasense--get-method-eldoc (method &optional index)
  (let* ((mtype (vbasense--method-type method))
         (name (vbasense--method-name method))
         (pinfos (vbasense--get-method-param-strings method))
         (rinfo (vbasense--get-method-ret-string method)))
    (concat (propertize mtype 'face 'font-lock-keyword-face)
            " "
            (propertize name 'face 'font-lock-function-name-face)
            "("
            (loop with i = 0
                  with ret = ""
                  for p in pinfos
                  if (> i 0)
                  do (setq ret (concat ret ", "))
                  if (and index
                          (= i index))
                  do (setq ret (concat ret (propertize p 'face 'bold)))
                  else
                  do (setq ret (concat ret p))
                  do (setq i (+ i 1))
                  finally return ret)
            ")" 
            rinfo)))


;;;;;;;;;;;;;;;;;;;;;;;
;; Other interactive

;;;###autoload
(defun vbasense-popup-help ()
  "Popup help about something at point."
  (interactive)
  (yaxception:$
    (yaxception:try
      (when (vbasense--active-p)
        (vbasense--update-current-instance)
        (let* ((x (vbasense--get-current-identifier))
               (doc (when x (vbasense--get-any-document :x x))))
          (cond ((not doc)
                 (message "[VBASense] Can't identify anything at point."))
                ((and (functionp 'ac-quick-help-use-pos-tip-p)
                      (ac-quick-help-use-pos-tip-p))
                 (pos-tip-show doc 'popup-tip-face nil nil 300 popup-tip-max-width))
                (t
                 (popup-tip doc))))))
    (yaxception:catch 'error e
      (message "[VBASense] %s" (yaxception:get-text e))
      (vbasense--error "failed popup help : %s\n%s"
                       (yaxception:get-text e)
                       (yaxception:get-stack-trace-string e)))))

;;;###autoload
(defun vbasense-jump-to-definition ()
  "Jump to definition at point."
  (interactive)
  (yaxception:$
    (yaxception:try
      (when (vbasense--active-p)
        (vbasense--update-current-instance)
        (let* ((x (vbasense--get-current-identifier))
               (fpath (cond ((vbasense--module-p x)     (vbasense--module-file x))
                            ((vbasense--interface-p x)  (vbasense--interface-file x))
                            ((vbasense--class-p x)      (vbasense--class-file x))
                            ((vbasense--enum-p x)       (vbasense--enum-file x))
                            ((vbasense--enummember-p x) (vbasense--enummember-file x))
                            ((vbasense--method-p x)
                             (let ((ns (gethash (vbasense--method-belongkey x) vbasense--hash-object-cache)))
                               (cond ((vbasense--module-p ns)    (vbasense--module-file ns))
                                     ((vbasense--interface-p ns) (vbasense--interface-file ns))
                                     ((vbasense--class-p ns)     (vbasense--class-file ns))
                                     ((member (vbasense--method-name x) vbasense--current-methods)
                                      (buffer-file-name)))))))
               (row (cond ((vbasense--class-p x)      (vbasense--class-line x))
                          ((vbasense--enum-p x)       (vbasense--enum-line x))
                          ((vbasense--enummember-p x) (vbasense--enummember-line x))
                          ((vbasense--method-p x)     (vbasense--method-line x)))))
          (if (or (not fpath)
                  (not (file-exists-p fpath)))
              (progn (message "[VBASense] Not found definition location at point.")
                     (vbasense--trace "Not found location fpath[%s] row[%s]" fpath row))
            (ring-insert find-tag-marker-ring (point-marker))
            (find-file fpath)
            (goto-char (point-min))
            (when row
              (forward-line (- row 1)))))))
    (yaxception:catch 'error e
      (message "[VBASense] %s" (yaxception:get-text e))
      (vbasense--error "failed jump to definition : %s\n%s"
                       (yaxception:get-text e)
                       (yaxception:get-stack-trace-string e)))))

(defvar vbasense--regexp-implements (rx-to-string `(and bol (* space)
                                                        (or ,@vbasense--words-implement) (+ space)
                                                        (group (regexp ,vbasense--regexp-namespace-ident)))))
(defun vbasense-insert-implement-definition ()
  "Insert definition of procedure which the implemented interface has."
  (interactive)
  (yaxception:$
    (yaxception:try
      (when (vbasense--active-p)
        (vbasense--trace "start insert implement definition")
        (let* ((ifs (save-excursion
                      (loop initially (goto-char (point-min))
                            while (re-search-forward vbasense--regexp-implements nil t)
                            for ifnm = (match-string-no-properties 1)
                            for belongkey = (concat "i:" (downcase ifnm))
                            for i = (gethash belongkey vbasense--hash-object-cache)
                            if (vbasense--interface-p i)
                            collect i
                            else
                            do (vbasense--info "can't detect implement object from %s" ifnm)))))
          (loop for i in ifs
                for ifnm = (vbasense--interface-name i)
                do (vbasense--trace "start insert definition of %s" ifnm)
                do (loop for m in (vbasense--interface-methods i)
                         for mtype = (vbasense--method-type m)
                         for mtdnm = (vbasense--method-name m)
                         for ptexts = (vbasense--get-method-param-strings m "Variant")
                         for rtext = (vbasense--get-method-ret-string m)
                         for deftext = (concat (format "Private %s %s_%s(%s)%s\n"
                                                       mtype ifnm mtdnm (mapconcat 'identity ptexts ", ") rtext)
                                               (format "End %s\n" mtype))
                         do (insert deftext "\n"))))
        (message "[VBASense] finished insert implement definition")))
    (yaxception:catch 'error e
      (message "[VBASense] %s" (yaxception:get-text e))
      (vbasense--error "failed insert implement definition : %s\n%s"
                       (yaxception:get-text e)
                       (yaxception:get-stack-trace-string e)))))


;;;;;;;;;;;;;;;
;; For setup

;;;###autoload
(defun vbasense-setup-current-buffer ()
  "Do setup for using vbasense in current buffer."
  (interactive)
  (yaxception:$
    (yaxception:try
      (vbasense--debug "start setup current buffer")
      (when (vbasense--active-p)
        ;; Key binding
        (loop for stroke in vbasense-ac-trigger-command-keys
              if (not (string= stroke ""))
              do (local-set-key (read-kbd-macro stroke) 'vbasense--insert-with-ac-trigger-command))
        (when (and (stringp vbasense-popup-help-key)
                   (not (string= vbasense-popup-help-key "")))
          (local-set-key (read-kbd-macro vbasense-popup-help-key) 'vbasense-popup-help))
        (when (and (stringp vbasense-jump-to-definition-key)
                   (not (string= vbasense-jump-to-definition-key "")))
          (local-set-key (read-kbd-macro vbasense-jump-to-definition-key) 'vbasense-jump-to-definition))
        ;; For auto-complete
        (add-to-list 'ac-sources 'ac-source-vbasense-head)
        (add-to-list 'ac-sources 'ac-source-vbasense-word)
        (add-to-list 'ac-sources 'ac-source-vbasense-operator)
        (add-to-list 'ac-sources 'ac-source-vbasense-dot)
        (add-to-list 'ac-sources 'ac-source-vbasense-paren)
        (add-to-list 'ac-sources 'ac-source-vbasense-comma)
        (add-to-list 'ac-sources 'ac-source-vbasense-keywordargument)
        (add-to-list 'ac-sources 'ac-source-vbasense-createobject)
        (add-to-list 'ac-sources 'ac-source-vbasense-implicit-enummember)
        (when (not ac-source-vbasense-implicit-enummember)
          (setq ac-source-vbasense-implicit-enummember (vbasense--get-ac-source-implicit-enummember)))
        (auto-complete-mode t)
        ;; For eldoc
        (set (make-local-variable 'eldoc-documentation-function) 'vbasense--echo-method-usage)
        (turn-on-eldoc-mode)
        (eldoc-add-command 'vbasense--insert-with-ac-trigger-command)
        ;; Other
        (vbasense-load-library)
        (vbasense-reload-current-buffer)
        (funcall vbasense-setup-user-library-function))
      (vbasense--debug "finish setup current buffer"))
    (yaxception:catch 'vbasense--ipc-err e
      (yaxception:throw 'vbasense--any-err
                        :func "setup current buffer"
                        :errmsg (yaxception:get-prop e 'errmsg)))
    (yaxception:catch 'vbasense--any-err e
      (yaxception:throw 'vbasense--any-err
                        :func "setup current buffer"
                        :errmsg (yaxception:get-prop e 'errmsg)))
    (yaxception:catch 'error e
      (vbasense--error "failed setup current buffer : %s\n%s"
                       (yaxception:get-text e) (yaxception:get-stack-trace-string e))
      (yaxception:throw 'vbasense--any-err
                        :func "setup current buffer"
                        :errmsg (yaxception:get-text e)))))

;;;###autoload
(defun vbasense-config-default ()
  "Do setting a recommemded configuration."
  (loop for tlipath in '("c:/WINDOWS/system32/wshom.ocx"
                         "c:/WINDOWS/system32/scrrun.dll"
                         "c:/WINDOWS/system32/vbscript.dll/3"
                         "c:/WINDOWS/system32/msxml.dll"
                         "c:/WINDOWS/system32/msxml2.dll")
        for fpath = (replace-regexp-in-string "/[0-9]+\\'" "" tlipath)
        if (file-exists-p fpath)
        do (add-to-list 'vbasense-tli-files tlipath))
  (loop for mode in vbasense-enable-modes
        for hook = (intern-soft (concat (symbol-name mode) "-hook"))
        do (add-to-list 'ac-modes mode)
        if (and hook
                (symbolp hook))
        do (add-hook hook 'vbasense-setup-current-buffer t))
  (add-hook 'after-save-hook 'vbasense-reload-current-buffer t))


(provide 'vbasense)
;;; vbasense.el ends here
