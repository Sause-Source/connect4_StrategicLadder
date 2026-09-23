# Connect4の戦略ラダー自動構築プログラム

盤面サイズと探索する手の深さを入力することで、
Connect4の戦略ラダーを構築、ゲームの深さ(depth)を自動算出できます。
出力はログを記録したcsvファイルと戦略ラダーを可視化したexcelファイルになっています。

resultsフォルダに実際に6×6,7×6,8×6の盤面サイズでシミュレーションを行った結果がまとめられています。
(ただし8×6は手の深さ37までの結果となっています。)

このプログラムはPascalPons氏のConnect4を基にして作成されています。
https://github.com/PascalPons/connect4

This C++ source code is published under AGPL v3 license.