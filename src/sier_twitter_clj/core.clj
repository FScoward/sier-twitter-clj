(ns sier-twitter-clj.core
  (:import [org.apache.poi.xssf.usermodel XSSFWorkbook]
           [java.io FileOutputStream]

           [twitter4j TwitterFactory Query TwitterException])
  (:require [clojure.java.io :as io]))

(def wb (XSSFWorkbook.))
(def twitter (. (TwitterFactory.) getInstance))

(defn save [content]
  (with-open [w (io/output-stream (str "外部インターフェース設計書_" (System/currentTimeMillis) ".xls"))]
    (.write content w)))

;; get twitter time line
;; return map
(defn get-timeline []
  (def x (.getHomeTimeline twitter))
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
    (let [cell (. row createCell 0)]
      (.setCellValue cell (str (:name tweet) "@" (:screenName tweet))))
    (let [cell (. row createCell 1)]
      (.setCellValue cell (:text tweet))))
  (save wb))
