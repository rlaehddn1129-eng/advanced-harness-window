#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""raffle.py 단위 테스트.  실행:  python3 -m unittest -v test_raffle.py"""

import contextlib
import io
import os
import shutil
import tempfile
import unittest
import zipfile
from collections import Counter

import raffle as R


def quiet(fn, *args, **kwargs):
    """CLI 실행 출력이 테스트 결과를 가리지 않도록 stdout을 삼킨다."""
    with contextlib.redirect_stdout(io.StringIO()):
        return fn(*args, **kwargs)


def entry(row_no, naming, name, email="", phone="", org="", dt=None):
    e = R.Entry(row_no=row_no, naming=naming, name=name,
                email=R.norm_email(email), phone=R.norm_phone(phone), org=org,
                submitted_at=R.parse_dt(dt) if dt else None)
    return e


class TestNormalize(unittest.TestCase):
    def test_phone_variants_collapse(self):
        forms = ["010-1234-5678", "01012345678", "010 1234 5678", "+82 10-1234-5678",
                 "+821012345678", "0082-10-1234-5678", "(010)1234-5678"]
        self.assertEqual({R.norm_phone(f) for f in forms}, {"01012345678"})

    def test_phone_rejects_garbage(self):
        for bad in ["", "번호없음", "123", "1" * 15]:
            self.assertEqual(R.norm_phone(bad), "")

    def test_email_alias_collapses(self):
        forms = ["hong@gmail.com", "HONG@Gmail.com", "h.o.n.g@gmail.com",
                 "hong+contest@gmail.com", "h.ong+1@googlemail.com"]
        self.assertEqual({R.norm_email(f) for f in forms}, {"hong@gmail.com"})

    def test_email_non_gmail_keeps_dots(self):
        self.assertEqual(R.norm_email("a.b@kwater.or.kr"), "a.b@kwater.or.kr")
        self.assertEqual(R.norm_email("a.b+x@kwater.or.kr"), "a.b@kwater.or.kr")

    def test_email_loose_mode_keeps_alias(self):
        self.assertNotEqual(R.norm_email("hong+1@gmail.com", strict=False),
                            R.norm_email("hong@gmail.com", strict=False))

    def test_invalid_email(self):
        for bad in ["", "hong", "hong@", "@gmail.com", "a@b@c.com", "hong@localhost"]:
            self.assertEqual(R.norm_email(bad), "")

    def test_display_text_keeps_original_jamo(self):
        self.assertEqual(R.display_text("  ㅋㅋㅋ  "), "ㅋㅋㅋ")


class TestIdentity(unittest.TestCase):
    def test_same_person_merged_by_phone_or_email(self):
        entries = [
            entry(2, "A", "박도윤", "doyoon@gmail.com", "010-0000-0001"),
            entry(3, "B", "박도윤", "d.o.yoon+1@gmail.com", "010-9999-9999"),  # 이메일로 연결
            entry(4, "C", "박 도윤", "other@example.com", "01000000001"),      # 전화로 연결
        ]
        persons = R.build_persons(entries)
        self.assertEqual(len(persons), 1)
        self.assertEqual(persons[0].submissions, 3)

    def test_homonyms_stay_separate(self):
        entries = [
            entry(2, "A", "김민준", "minjun1@example.com", "010-0000-0001", "K-water"),
            entry(3, "B", "김민준", "minjun2@example.com", "010-0000-0002", "고려대"),
        ]
        persons = R.build_persons(entries)
        self.assertEqual(len(persons), 2, "동명이인은 별개 인원이어야 한다")
        self.assertEqual(len(R.duplicate_name_groups(persons)), 1)

    def test_ten_submissions_get_one_ticket(self):
        entries = [entry(i, f"안{i}", "박도윤", "doyoon@gmail.com", "010-0000-0001")
                   for i in range(2, 12)]
        entries.append(entry(12, "정상", "이서연", "seo@example.com", "010-0000-0002"))
        persons = R.build_persons(entries)
        self.assertEqual(len(persons), 2)
        self.assertEqual({p.submissions for p in persons}, {10, 1})

    def test_identity_key_is_order_independent(self):
        a = R.build_persons([entry(2, "A", "홍", "h@example.com", "010-0000-0001"),
                             entry(3, "B", "홍", "h2@example.com", "010-0000-0001")])
        b = R.build_persons([entry(2, "B", "홍", "h2@example.com", "010-0000-0001"),
                             entry(3, "A", "홍", "h@example.com", "010-0000-0001")])
        self.assertEqual(a[0].identity_key, b[0].identity_key)

    def test_representative_entry_is_earliest_submission(self):
        persons = R.build_persons([
            entry(2, "나중", "홍", "h@example.com", "010-0000-0001", dt="2026-09-05 10:00"),
            entry(3, "처음", "홍", "h@example.com", "010-0000-0001", dt="2026-08-28 09:00"),
        ])
        self.assertEqual(persons[0].primary.naming, "처음")

    def test_entry_without_contact_is_invalid(self):
        rows = [{"네이밍": "A", "성명": "홍길동", "이메일": "", "연락처": ""}]
        cols = R.detect_columns(list(rows[0]), {})
        entries = R.build_entries(rows, cols)
        self.assertFalse(entries[0].valid)
        self.assertEqual(R.build_persons(entries), [])


class TestFlags(unittest.TestCase):
    def test_junk_naming_flagged(self):
        self.assertTrue(R.naming_flags("ㅋㅋㅋ"))
        self.assertTrue(R.naming_flags("ㅁㄴㅇㄹ"))
        self.assertTrue(R.naming_flags("aaaa"))
        self.assertTrue(R.naming_flags("!!!"))
        self.assertTrue(R.naming_flags("A"))
        self.assertTrue(R.naming_flags("홍길동", person_name="홍길동"))

    def test_normal_naming_not_flagged(self):
        for good in ["AquaLink", "물길이음", "GARAM ONE", "Water+"]:
            self.assertEqual(R.naming_flags(good), [], good)


class TestDraw(unittest.TestCase):
    def setUp(self):
        self.persons = R.build_persons([
            entry(i, f"안{i}", f"응모자{i}", f"user{i}@example.com", f"010-0000-{i:04d}")
            for i in range(2, 62)
        ])

    def test_deterministic(self):
        a, _ = R.draw_winners(self.persons, "seed-1", 10)
        b, _ = R.draw_winners(self.persons, "seed-1", 10)
        self.assertEqual([p.identity_key for p in a], [p.identity_key for p in b])

    def test_independent_of_roster_order(self):
        a, _ = R.draw_winners(self.persons, "seed-1", 10)
        b, _ = R.draw_winners(list(reversed(self.persons)), "seed-1", 10)
        self.assertEqual([p.identity_key for p in a], [p.identity_key for p in b])

    def test_seed_changes_result(self):
        a, _ = R.draw_winners(self.persons, "seed-1", 10)
        b, _ = R.draw_winners(self.persons, "seed-2", 10)
        self.assertNotEqual([p.identity_key for p in a], [p.identity_key for p in b])

    def test_winners_and_reserve_disjoint(self):
        won, reserve = R.draw_winners(self.persons, "seed-1", 10, 5)
        self.assertEqual(len(won), 10)
        self.assertEqual(len(reserve), 5)
        self.assertFalse({p.pid for p in won} & {p.pid for p in reserve})

    def test_roster_hash_ignores_order(self):
        self.assertEqual(R.roster_hash(self.persons), R.roster_hash(list(reversed(self.persons))))

    def test_roster_hash_changes_when_member_changes(self):
        self.assertNotEqual(R.roster_hash(self.persons), R.roster_hash(self.persons[:-1]))

    def test_uniform_enough(self):
        """시드 2000회 반복 시 특정인에게 쏠리지 않는지(기대 당첨률 10/60 ≈ 16.7%)."""
        trials, counts = 2000, Counter()
        for i in range(trials):
            won, _ = R.draw_winners(self.persons, f"seed-{i}", 10)
            counts.update(p.pid for p in won)
        rates = [counts[p.pid] / trials for p in self.persons]
        self.assertAlmostEqual(sum(rates) / len(rates), 10 / 60, places=6)
        self.assertLess(max(rates), 0.22, "특정 응모자에게 유리하면 안 된다")
        self.assertGreater(min(rates), 0.11, "특정 응모자에게 불리하면 안 된다")


class TestExclusion(unittest.TestCase):
    def setUp(self):
        self.persons = R.build_persons([
            entry(2, "A", "홍길동", "hong@example.com", "010-0000-0001", "K-water"),
            entry(3, "B", "홍길동", "hong2@example.com", "010-0000-0002", "부산대"),
            entry(4, "C", "이순신", "lee@example.com", "010-0000-0003"),
            entry(5, "D", "이순신", "lee@example.com", "010-0000-0003"),
        ])

    def test_exclude_by_email_or_phone(self):
        for token in ("hong@example.com", "010-0000-0001", "01000000001"):
            resolved = R.resolve_exclusions(self.persons, [token], True, "테스트")
            self.assertEqual(len(resolved), 1)
            self.assertEqual(resolved[0][1].org, "K-water")

    def test_exclude_removes_all_entries_of_that_person(self):
        """최상위 1건만 시상: 수상자가 낸 다른 제출건도 함께 빠져야 한다."""
        resolved = R.resolve_exclusions(self.persons, ["lee@example.com"], True, "테스트")
        excluded = resolved[0][1]
        self.assertEqual(excluded.submissions, 2)
        roster = [p for p in self.persons if p.pid != excluded.pid]
        self.assertEqual(sum(p.submissions for p in roster), 2)

    def test_ambiguous_name_is_rejected(self):
        with self.assertRaises(SystemExit):
            R.resolve_exclusions(self.persons, ["홍길동"], True, "테스트")

    def test_unknown_token_is_rejected(self):
        with self.assertRaises(SystemExit):
            R.resolve_exclusions(self.persons, ["nobody@example.com"], True, "테스트")

    def test_unique_name_is_accepted(self):
        resolved = R.resolve_exclusions(self.persons, ["이순신"], True, "테스트")
        self.assertEqual(len(resolved), 1)


class TestColumnDetection(unittest.TestCase):
    def test_naming_wins_over_name(self):
        cols = R.detect_columns(["제출일시", "네이밍(제안명)", "성명", "이메일", "연락처", "소속"], {})
        self.assertEqual(cols["naming"], "네이밍(제안명)")
        self.assertEqual(cols["name"], "성명")
        self.assertEqual(cols["phone"], "연락처")

    def test_english_headers(self):
        cols = R.detect_columns(["Timestamp", "Proposal", "Name", "Email", "Phone"], {})
        self.assertEqual(cols["naming"], "Proposal")
        self.assertEqual(cols["email"], "Email")

    def test_override(self):
        cols = R.detect_columns(["가", "나", "다"], {"naming": "가", "name": "나"})
        self.assertEqual((cols["naming"], cols["name"]), ("가", "나"))

    def test_override_rejects_missing_column(self):
        with self.assertRaises(SystemExit):
            R.detect_columns(["가"], {"naming": "없는열"})


def make_xlsx(path, rows):
    """openpyxl 없이 테스트용 xlsx를 만든다(공유문자열 미사용, inlineStr)."""
    def cell(ci, value):
        ref = ""
        n = ci + 1
        while n:
            n, rem = divmod(n - 1, 26)
            ref = chr(65 + rem) + ref
        esc = (str(value).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
        return f'<c r="{ref}{"{r}"}" t="inlineStr"><is><t>{esc}</t></is></c>'

    body = []
    for ri, row in enumerate(rows, start=1):
        cells = "".join(cell(ci, v).replace("{r}", str(ri)) for ci, v in enumerate(row))
        body.append(f'<row r="{ri}">{cells}</row>')
    sheet = ('<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.'
             'openxmlformats.org/spreadsheetml/2006/main"><sheetData>'
             + "".join(body) + "</sheetData></worksheet>")
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("[Content_Types].xml",
                    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/'
                    'package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/>'
                    "</Types>")
        zf.writestr("xl/worksheets/sheet1.xml", sheet)


class TestXlsxReader(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_reads_shared_strings_and_named_sheet(self):
        path = os.path.join(self.tmp, "real.xlsx")
        make_xlsx_shared(path, [["제출일시", "네이밍", "성명", "이메일", "연락처"],
                                ["2026-08-28 09:00", "AquaLink", "홍길동", "hong@gmail.com", "010-1111-2222"],
                                ["2026-08-28 09:10", "WaterOne", "홍길동", "h.ong+1@gmail.com", "010-1111-2222"],
                                ["2026-08-29 10:00", "PureFlow", "이순신", "lee@example.com", "010-3333-4444"]])
        rows = R.read_table(path)
        self.assertEqual(len(rows), 3)
        cols = R.detect_columns(list(rows[0]), {})
        persons = R.build_persons(R.build_entries(rows, cols))
        self.assertEqual(len(persons), 2, "별칭 이메일 2건은 한 사람으로 합쳐져야 한다")

    def test_reads_simple_sheet(self):
        path = os.path.join(self.tmp, "t.xlsx")
        make_xlsx(path, [["네이밍", "성명", "이메일", "연락처"],
                         ["AquaLink", "홍길동", "hong@example.com", "010-1111-2222"],
                         ["WaterOne", "이순신", "lee@example.com", "010-3333-4444"]])
        rows = R.read_table(path)
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0]["네이밍"], "AquaLink")
        cols = R.detect_columns(list(rows[0]), {})
        persons = R.build_persons(R.build_entries(rows, cols))
        self.assertEqual(len(persons), 2)


def make_xlsx_shared(path, rows, sheet_name="worksheets/데이터.xml"):
    """Excel이 실제로 쓰는 구조: sharedStrings + workbook rels + sheet1 아닌 파일명."""
    strings, index = [], {}

    def sid(value):
        if value not in index:
            index[value] = len(strings)
            strings.append(value)
        return index[value]

    def ref(ci, ri):
        s, n = "", ci + 1
        while n:
            n, rem = divmod(n - 1, 26)
            s = chr(65 + rem) + s
        return f"{s}{ri}"

    body = []
    for ri, row in enumerate(rows, start=1):
        cells = "".join(f'<c r="{ref(ci, ri)}" t="s"><v>{sid(v)}</v></c>'
                        for ci, v in enumerate(row) if v)
        body.append(f'<row r="{ri}">{cells}</row>')
    ns = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("xl/sharedStrings.xml", f"<sst {ns}>"
                    + "".join(f"<si><t>{s}</t></si>" for s in strings) + "</sst>")
        zf.writestr("xl/workbook.xml", f'<workbook {ns} xmlns:r="http://schemas.'
                    'openxmlformats.org/officeDocument/2006/relationships"><sheets>'
                    '<sheet name="응모현황" sheetId="1" r:id="rId1"/></sheets></workbook>')
        zf.writestr("xl/_rels/workbook.xml.rels", '<Relationships xmlns="http://schemas.'
                    'openxmlformats.org/package/2006/relationships">'
                    f'<Relationship Id="rId1" Target="{sheet_name}"/></Relationships>')
        zf.writestr(f"xl/{sheet_name}", f"<worksheet {ns}><sheetData>"
                    + "".join(body) + "</sheetData></worksheet>")


class TestEndToEnd(unittest.TestCase):
    """실제 CLI 흐름: draw → verify 가 재현되는지."""

    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.csv = os.path.join(self.tmp, "응모.csv")
        lines = ["제출일시,네이밍,성명,이메일,연락처,소속"]
        for i in range(1, 41):
            lines.append(f"2026-08-{28 + i % 3:02d} 09:00,안{i},응모자{i},"
                         f"user{i}@example.com,010-5555-{i:04d},소속{i % 5}")
        for k in range(10):  # 다량 제출자 1명
            lines.append(f"2026-09-01 10:{k:02d},스팸{k},스패머,spam+{k}@gmail.com,010-7777-0001,")
        with open(self.csv, "w", encoding="utf-8") as fp:
            fp.write("\n".join(lines) + "\n")
        self.out = os.path.join(self.tmp, "결과")

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def _draw(self, extra=()):
        argv = ["draw", "--input", self.csv, "--winners", "10", "--reserve", "3",
                "--seed", "테스트 시드", "--out", self.out, *extra]
        return quiet(R.main, argv)

    def test_draw_then_verify(self):
        self.assertEqual(self._draw(["--exclude-winner", "user1@example.com",
                                     "--exclude-winner", "010-5555-0002",
                                     "--exclude-winner", "응모자3"]), 0)
        for name in ("01_당첨자.csv", "02_추첨대상_명부.csv", "03_제외자.csv",
                     "04_무효응모.csv", "05_확인필요_제안.csv", "audit.json", "report.md"):
            self.assertTrue(os.path.exists(os.path.join(self.out, name)), name)
        self.assertEqual(quiet(R.main, ["verify", "--input", self.csv,
                                        "--audit", os.path.join(self.out, "audit.json")]), 0)

    def test_verify_fails_after_roster_change(self):
        self.assertEqual(self._draw(), 0)
        with open(self.csv, "a", encoding="utf-8") as fp:
            fp.write("2026-09-11 10:00,추가,신규응모자,new@example.com,010-8888-0001,\n")
        self.assertEqual(quiet(R.main, ["verify", "--input", self.csv,
                                        "--audit", os.path.join(self.out, "audit.json")]), 1)

    def test_more_than_three_top_winners_rejected(self):
        with self.assertRaises(SystemExit):
            self._draw(["--exclude-winner", "user1@example.com",
                        "--exclude-winner", "user2@example.com",
                        "--exclude-winner", "user3@example.com",
                        "--exclude-winner", "user4@example.com"])

    def test_spammer_counted_once(self):
        self._draw()
        with open(os.path.join(self.out, "02_추첨대상_명부.csv"), encoding="utf-8-sig") as fp:
            body = fp.read()
        self.assertEqual(body.count("스패머"), 1, "10건을 낸 응모자도 명부에는 1행만 있어야 한다")

    def test_not_enough_candidates(self):
        with self.assertRaises(SystemExit):
            quiet(R.main, ["draw", "--input", self.csv, "--winners", "999",
                           "--seed", "s", "--out", self.out])


if __name__ == "__main__":
    unittest.main(verbosity=2)
