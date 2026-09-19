import unittest as ut

import Torob_Crawling


class TestDivide(ut.TestCase):
    def test_divide(self):
        self.assertEqual(
            Torob_Crawling.TorobScraperApp.price_data_cleaner("۶۳٫۹۹۰٫۰۰۰ تومان"),
            (63990000, "موجود"),
        )


if __name__ == "__main__":
    ut.main()
"""

from re import findall as fa


def price_data_cleaner(raw_price):

    price_number_str = raw_price.replace(",", "").replace("،", "").replace("٫", "")
    status = "نا موجود" if "نا موجود" in raw_price else "موجود"
    price_number = fa(r"\d+", price_number_str)
    if price_number:
        return int(price_number[0]), status
    return None, status


print(price_data_cleaner("۶۳٫۹۹۰٫۰۰۰ تومان"))
"""
