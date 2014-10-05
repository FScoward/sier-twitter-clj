(ns sier-twitter-clj.core
  (:import [org.apache.poi.xssf.usermodel XSSFWorkbook]
           [org.apache.poi.ss.usermodel IndexedColors CellStyle]
           [java.io FileOutputStream]
           [twitter4j TwitterFactory Query TwitterException Paging])
  (:require [clojure.java.io :as io]))

(def wb (XSSFWorkbook.))
(def twitter (. (TwitterFactory.) getInstance))

(defn save [content]
  (with-open [w (io/output-stream (str "外部インターフェース設計書_" (System/currentTimeMillis) ".xlsx"))]
    (.write content w)))

;; get twitter time line
;; return map
(defn get-timeline []
  (def x (.getHomeTimeline twitter (Paging. (int 1) (int 200))))
  (map #(zipmap [:screenName :name :text] [(.. % getUser getScreenName)
                                           (.. % getUser getName)
                                           (.getText %)
                                           ]) x))

(defn -main []
  (def timeline (get-timeline))
  (def tweets-size (count timeline))
  (def sheet (.createSheet wb))
  (dotimes [i tweets-size]
    (def tweet (first (drop (- tweets-size (inc i)) timeline)))
    (def row (.createRow sheet i))
    (def style (.createCellStyle wb))
    (.setFillForegroundColor style (.. IndexedColors AQUA getIndex))
    (.setFillPattern style (CellStyle/SOLID_FOREGROUND))
    (let [cell (. row createCell 0)]
      (.setCellValue cell (str (:name tweet) "@" (:screenName tweet)))
      (if (= 0 (rem i 2))
        (.setCellStyle cell style)))
    (let [cell (. row createCell 1)]
      (.setCellValue cell (:text tweet))
      (if (= 0 (rem i 2))
        (.setCellStyle cell style))))
  (save wb))
