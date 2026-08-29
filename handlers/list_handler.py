"""
ListHandler: Pythonのリストと同じ0始まりの番号で扱える「リスト」を
組み立てる/操作するためのハンドラ。

Excelのセル範囲を読み込む get_range_as_list(ExcelHandler側)や、
このハンドラの append で組み立てたリストは、store_as で変数に保存すれば、
後の手順で {{変数名[0]}} のように0始まりの番号で直接読み出せる
(この構文はExecutorの変数置換機構が解釈するため、このハンドラ自身は
インデックスアクセスの機能を持たない。get_item/lengthは、テンプレートの
外で明示的に1手順として使いたい場合の代替手段として用意している)。

for_start/for_end(制御構文)と組み合わせることで、VBAの
「For i = 0 To リストの要素数-1 ... Next」に相当する繰り返し処理ができる。
"""
from __future__ import annotations

from typing import Any


class ListHandler:
    def create_empty(self) -> list[Any]:
        """空のリストを作る(この後 append で1件ずつ積み上げていく用途)。"""
        return []

    def from_values(self, values: list[Any]) -> list[Any]:
        """カンマ区切り等で用意した値の並びを、そのままリストとして取り込む。"""
        return list(values)

    def append(self, lst: list[Any], value: Any) -> list[Any]:
        """リストの末尾に値を追加した、新しいリストを返す。
        store_as で元の変数名と同じ名前に保存し直せば、
        「そのリストに追記していく」という使い方になる。
        """
        if not isinstance(lst, list):
            raise ValueError(f"appendの対象はリストである必要があります(実際: {type(lst).__name__})")
        return lst + [value]

    def get_item(self, lst: list[Any], index: int) -> Any:
        """リストの指定位置の要素を取得する(0始まり)。
        通常は {{変数名[0]}} のようにテンプレート内で直接書く方が簡単だが、
        取得した値をさらに別の変数として保存し直したい場合等に使う。
        """
        if not isinstance(lst, list):
            raise ValueError(f"get_itemの対象はリストである必要があります(実際: {type(lst).__name__})")
        idx = int(index)
        if idx < 0 or idx >= len(lst):
            raise IndexError(f"範囲外のインデックスです: {idx}(要素数{len(lst)}、0〜{len(lst) - 1}の範囲で指定)")
        return lst[idx]

    def length(self, lst: list[Any]) -> int:
        """リストの要素数を取得する。通常は {{変数名.length}} で直接書ける。"""
        if not isinstance(lst, list):
            raise ValueError(f"lengthの対象はリストである必要があります(実際: {type(lst).__name__})")
        return len(lst)
