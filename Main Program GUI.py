import pandas as pd
import random
import os
from collections import defaultdict, Counter
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
import pandas as pd
import threading
import tkinter as tk
from ttkbootstrap.widgets import Separator
import sys

# ==========================
# إعداد البيانات
# ==========================
# (سيتم التحميل لاحقًا عند الضغط على "تشغيل الجدولة" وفق LAZY_INIT)
professors_df = pd.DataFrame()
rooms_df = pd.DataFrame()
terms_df = pd.DataFrame()

# تنظيف أسماء المواد سيتم عند التهيئة اللاحقة

# ==========================
# إعدادات عامة قابلة للضبط
# ==========================
random.seed(42)

# 1 = توزيع حسب عدد الطلاب، 2 = توزيع حسب عدد الشعب
DISTRIBUTION_MODE = 2

# اختيار الفصل: 1 أو 2. ضع None لإلغاء الفلترة وأخذ كل الفصول.
SEMESTER_FILTER = 2  # مثال: اختر الفصل الثاني. غيّره إلى 1 أو None حسب الحاجة.

# ساعات كل نوع
HOURS_THEORY = 3.0
HOURS_LAB = 1.5

# الحدود الدنيا لفتح شعبة (عند التوزيع بالطلاب)
MIN_THEORY = 20
MIN_LAB = 10

# إعدادات GA
POP_SIZE = 80
GENERATIONS = 220
ELITE = 8
TOURNAMENT_K = 3
MUT_PROB = 0.18
EARLY_STOP_NO_IMPROVE = 60   # إيقاف مبكر إذا لم يتحسن الأفضل خلال N أجيال

# أوزان العقوبات/المكافآت
PENALTY_ROOM_CONFLICT = 800
PENALTY_PROF_CONFLICT = 800
PENALTY_NO_ROOM = 100
PENALTY_YEAR_HARD = 200   # تضارب سنة لمادتين كلٌ منهما شعبة واحدة
PENALTY_YEAR_SOFT = 40    # تضارب سنة بوجود بدائل
GAP_WEIGHT = 8            # فجوات بين محاضرات نفس اليوم
VARIANCE_WEIGHT = 2.0     # عدالة توزيع الحمل
ASSIGNED_BONUS = 90       # مكافأة على كل شعبة مُسنَدة
OVER_LIMIT_HARD = 100 # عقوبة كبيرة جدًا لو حدث تجاوز ساعات (احتياطيًا)

# محاولات البحث عن مكان/قاعة متاحة لكل شعبة
MAX_ASSIGN_TRIES = 40

# ==========================
# الزمن والأيام
# ==========================
TIME_MAP = {
    "8-9": 1, "9-10": 2, "10-11": 3, "11-12": 4, "12-13": 5, "13-14": 6,
    "8-9:30": 1.5, "9:30-11": 2.5, "11-12:30": 3.5, "12:30-14": 4.5,
    "8-11": 1, "11-14": 2, "14-17": 3
}

THEORY_DAYS_SETS = [["ح", "ث", "خ"], ["ن", "ر"]]
THEORY_SHORT_SLOTS = ["8-9", "9-10", "10-11", "11-12", "12-13", "13-14"]
THEORY_LONG_SLOTS  = ["8-9:30", "9:30-11", "11-12:30", "12:30-14"]
LAB_DAYS = ["ح", "ن", "ث", "ر", "خ"]
LAB_SLOTS = ["8-11", "11-14", "14-17"]

# ==========================
# أدوات مساعدة وقراءة التخصصات (عام فقط)
# ==========================
def parse_courses(row):
    return [c.strip().replace('"', '') for c in str(row).split(",") if str(c).strip()]

def course_kind(name: str) -> str:
    return "عملي" if "مختبر" in name else "نظري"

def split_multi(val):
    if pd.isna(val):
        return []
    s = str(val).replace("،", ",")
    return [t.strip() for t in s.split(',') if t.strip()]

def parse_semesters(val):
    """يرجع مجموعة {1,2} حسب قيمة عمود 'الفصل' مثل '1', '2', '1,2' (يدعم الفاصلة العربية)."""
    if pd.isna(val):
        return set()
    s = str(val).strip().replace("،", ",")
    parts = [p.strip() for p in s.split(",") if p.strip()]
    out = set()
    for p in parts:
        if p in ("1", "٢"):
            out.add(1)
        elif p in ("2", "٢"):
            out.add(2)
        else:
            try:
                n = int(p)
                if n in (1, 2):
                    out.add(n)
            except:
                pass
    return out

def matches_semester(row):
    """فلترة الصف حسب SEMESTER_FILTER: يقبل المادة إذا احتوت مجموعة الفصول على الاختيار."""
    if SEMESTER_FILTER is None:
        return True
    sems = parse_semesters(row.get("الفصل", ""))
    return SEMESTER_FILTER in sems

# ==========================
# تفعيل التهيئة الكسولة (GUI أولاً)
# ==========================
LAZY_INIT = True

# ==========================
# بيانات الأساتذة/القاعات/المواد (ستُبنى لاحقًا إن LAZY_INIT=True)
# ==========================
PROF_COURSES = {}
PROF_LIMITS = {}
PROF_GENERAL = {}

ROOMS_BY_KIND = {"نظري": [], "عملي": []}
ROOM_CAP = {}

# تخصصات المواد (عام: قائمة) — من الجدول المُفلتر
COURSE_GENERAL = {}

# بناء الشعب المطلوبة (من الجدول المُفلتر)
SECTIONS = []  # كل عنصر = {course, kind, year, size}
COURSE_SECTIONS_COUNT = Counter()
COURSE_YEAR = {}
COURSE_KIND = {}

# إن أردت التهيئة المبكرة (غير مفعلة هنا)
if not LAZY_INIT:
    # هنا توضع قراءة الإكسل + الفلترة + البناء (مُعطّل الآن)
    pass

# ==========================
# اختيار الأساتذة: مرحلتان (صارم ثم عام)
# ==========================
def hours_of_kind(kind: str) -> float:
    return HOURS_LAB if kind == "عملي" else HOURS_THEORY

def eligible_profs(course: str, used_hours: dict, h_needed: float, mode: str):
    cand = []
    if mode == 'strict':
        for p in PROF_LIMITS.keys():
            if used_hours[p] + h_needed <= PROF_LIMITS.get(p, 0.0):
                if course in PROF_COURSES.get(p, set()):
                    cand.append(p)
    else:  # general
        gen_list = COURSE_GENERAL.get(course, [])
        gen_set = set(gen_list) if gen_list else set()
        if gen_set:
            for p in PROF_LIMITS.keys():
                if used_hours[p] + h_needed <= PROF_LIMITS.get(p, 0.0):
                    if PROF_GENERAL.get(p, set()).intersection(gen_set):
                        cand.append(p)
    cand.sort(key=lambda x: used_hours[x])
    return cand

def pick_prof_for_section_strict(sec, used_hours: dict):
    h = hours_of_kind(sec["kind"])
    cands = eligible_profs(sec["course"], used_hours, h, mode='strict')
    return cands[0] if cands else None

def pick_prof_for_section_general(sec, used_hours: dict):
    h = hours_of_kind(sec["kind"])
    cands = eligible_profs(sec["course"], used_hours, h, mode='general')
    return cands[0] if cands else None

# ==========================
# توافر القاعة/الأستاذ + بناء إشغال
# ==========================
def rebuild_occupancies(chrom):
    room_occ = set()               # (room, day, slot)
    prof_occ = defaultdict(set)    # prof -> {(day, slot)}
    year_occ_single = defaultdict(set)  # year -> {(day, slot)} لمواد single-section فقط
    for g in chrom:
        if g is None:
            continue
        r = g["room"]
        p = g["prof"]
        slot = g["slot"]
        y = g["year"]
        c = g["course"]
        for d in g["days"]:
            room_occ.add((r, d, slot))
            prof_occ[p].add((d, slot))
            if COURSE_SECTIONS_COUNT.get(c, 0) == 1:
                year_occ_single[y].add((d, slot))
    return room_occ, prof_occ, year_occ_single

def year_single_conflict(sec, day, slot, year_occ_single):
    """لا نسمح بتضارب إذا كانت المادة الحالية single-section ويتصادم وقتها مع single-section أخرى بنفس السنة."""
    c = sec["course"]
    y = sec["year"]
    if COURSE_SECTIONS_COUNT.get(c, 0) == 1:
        return (day, slot) in year_occ_single[y]
    return False

def find_feasible_assignment_for_prof(sec, prof, room_occ, prof_occ, year_occ_single):
    """يحاول إيجاد (days, slot, room) بلا تعارض للقاعة/الأستاذ وقيود السنة."""
    kind = sec["kind"]
    if kind == "نظري":
        patterns = []
        patterns.append((THEORY_DAYS_SETS[0][:], THEORY_SHORT_SLOTS))
        patterns.append((THEORY_DAYS_SETS[1][:], THEORY_LONG_SLOTS))
        random.shuffle(patterns)
        for _ in range(MAX_ASSIGN_TRIES):
            days_set, slots = random.choice(patterns)
            slot = random.choice(slots)
            rooms = ROOMS_BY_KIND.get(kind, [])
            random.shuffle(rooms)

            # تحقق توافر الأستاذ والسنة
            conflict = False
            for d in days_set:
                if (d, slot) in prof_occ[prof]:
                    conflict = True
                    break
                if year_single_conflict(sec, d, slot, year_occ_single):
                    conflict = True
                    break
            if conflict:
                continue

            ok_room = None
            for r in rooms:
                if any(((r, d, slot) in room_occ) for d in days_set):
                    continue
                ok_room = r
                break
            if ok_room is None:
                continue
            return days_set, slot, ok_room
        return None
    else:
        for _ in range(MAX_ASSIGN_TRIES):
            d = random.choice(LAB_DAYS)
            slot = random.choice(LAB_SLOTS)
            if (d, slot) in prof_occ[prof]:
                continue
            if year_single_conflict(sec, d, slot, year_occ_single):
                continue
            rooms = ROOMS_BY_KIND.get(kind, [])
            random.shuffle(rooms)
            ok_room = None
            for r in rooms:
                if (r, d, slot) in room_occ:
                    continue
                ok_room = r
                break
            if ok_room is None:
                continue
            return [d], slot, ok_room
        return None

def apply_assignment_to_occ(prof, days, slot, room, sec, room_occ, prof_occ, year_occ_single):
    for d in days:
        room_occ.add((room, d, slot))
        prof_occ[prof].add((d, slot))
        if COURSE_SECTIONS_COUNT.get(sec["course"], 0) == 1:
            year_occ_single[sec["year"]].add((d, slot))

# ==========================
# توليد كروموسوم (قابل) + إصلاح صارم
# ==========================
def compute_used_hours(chrom):
    used = defaultdict(float)
    for g in chrom:
        if g is None:
            continue
        p = g.get("prof")
        if p is None:
            continue
        used[p] += HOURS_LAB if g["kind"] == "عملي" else HOURS_THEORY
    return used

def random_chromosome():
    chrom = []
    used_hours = defaultdict(float)
    room_occ, prof_occ, year_occ_single = set(), defaultdict(set), defaultdict(set)
    pending = []  # غير المُسنَّد في المرحلة 1

    # المرحلة 1: صارم (المواد فقط)
    for sec in SECTIONS:
        prof = pick_prof_for_section_strict(sec, used_hours)
        if prof is None:
            pending.append(sec)
            continue
        found = find_feasible_assignment_for_prof(sec, prof, room_occ, prof_occ, year_occ_single)
        if not found:
            pending.append(sec)
            continue
        days, slot, room = found
        chrom.append({
            "course": sec["course"],
            "kind": sec["kind"],
            "year": sec["year"],
            "size": sec["size"],
            "prof": prof,
            "room": room,
            "days": days,
            "slot": slot,
        })
        used_hours[prof] += HOURS_LAB if sec["kind"] == "عملي" else HOURS_THEORY
        apply_assignment_to_occ(prof, days, slot, room, sec, room_occ, prof_occ, year_occ_single)

    # المرحلة 2: عام (لغير المسند)
    for sec in pending:
        prof = pick_prof_for_section_general(sec, used_hours)
        if prof is None:
            continue
        found = find_feasible_assignment_for_prof(sec, prof, room_occ, prof_occ, year_occ_single)
        if not found:
            continue
        days, slot, room = found
        chrom.append({
            "course": sec["course"],
            "kind": sec["kind"],
            "year": sec["year"],
            "size": sec["size"],
            "prof": prof,
            "room": room,
            "days": days,
            "slot": slot,
        })
        used_hours[prof] += HOURS_LAB if sec["kind"] == "عملي" else HOURS_THEORY
        apply_assignment_to_occ(prof, days, slot, room, sec, room_occ, prof_occ, year_occ_single)

    return chrom

def try_reassign_gene(chrom, idx):
    """يحاول إعادة تعيين جين مع احترام القيود الصلبة (غرفة/أستاذ/سنة-واحدة)."""
    g = chrom[idx]
    tmp = chrom[:idx] + chrom[idx+1:]
    room_occ, prof_occ, year_occ_single = rebuild_occupancies(tmp)
    used = compute_used_hours(tmp)

    sec = {"course": g["course"], "kind": g["kind"], "year": g["year"], "size": g["size"]}
    hours = HOURS_LAB if g["kind"] == "عملي" else HOURS_THEORY

    candidates = [g["prof"]]
    strict = eligible_profs(g["course"], used, hours, mode='strict')
    general = eligible_profs(g["course"], used, hours, mode='general')
    for lst in (strict, general):
        for p in lst:
            if p not in candidates:
                candidates.append(p)

    random.shuffle(candidates)

    for p in candidates:
        found = find_feasible_assignment_for_prof(sec, p, room_occ, prof_occ, year_occ_single)
        if found:
            days, slot, room = found
            chrom[idx] = {
                "course": sec["course"],
                "kind": sec["kind"],
                "year": sec["year"],
                "size": sec["size"],
                "prof": p,
                "room": room,
                "days": days,
                "slot": slot,
            }
            return True
    return False

def repair_chromosome(chrom, max_passes=200):
    """إصلاح صارم لتجاوز الساعات + محاولة حل التعارضات عبر إعادة التعيين."""
    chrom = [g for g in chrom if g is not None]
    passes = 0
    changed = True

    while changed and passes < max_passes:
        passes += 1
        changed = False
        used = compute_used_hours(chrom)

        # 1) لا تتجاوز حد الساعات
        over = [p for p, h in used.items() if h > PROF_LIMITS.get(p, 0.0) + 1e-9]
        if over:
            for p in over:
                idxs = [i for i, g in enumerate(chrom) if g is not None and g.get("prof") == p]
                for i in idxs:
                    ok = try_reassign_gene(chrom, i)
                    if ok:
                        changed = True
                    else:
                        chrom[i] = None
                        changed = True
                    used = compute_used_hours([g for g in chrom if g is not None])
                    if used.get(p, 0.0) <= PROF_LIMITS.get(p, 0.0) + 1e-9:
                        break
            chrom = [g for g in chrom if g is not None]
            continue

        # 2) إصلاح أي تضارب زمني محتمل
        room_occ, prof_occ, year_occ_single = rebuild_occupancies(chrom)
        for i, g in enumerate(list(chrom)):
            prof = g["prof"]
            bad = False
            for d in g["days"]:
                if list(prof_occ[prof]).count((d, g["slot"])) > 1:
                    bad = True
                    break
            if bad:
                ok = try_reassign_gene(chrom, i)
                if ok:
                    changed = True

        chrom = [g for g in chrom if g is not None]

    return chrom

# ==========================
# دالة اللياقة
# ==========================
def fitness(chrom):
    chrom = [g for g in chrom if g is not None]

    score = 0.0
    penalty = 0.0

    room_time = set()
    prof_time = set()
    year_time_courses = defaultdict(lambda: defaultdict(set))

    used_hours = compute_used_hours(chrom)

    for g in chrom:
        c = g["course"]
        k = g["kind"]
        y = g["year"]
        p = g["prof"]
        r = g["room"]
        days = g["days"]
        slot = g["slot"]

        if r is None:
            penalty += PENALTY_NO_ROOM

        for d in days:
            kr = (r, d, slot)
            if r is None:
                pass
            elif kr in room_time:
                penalty += PENALTY_ROOM_CONFLICT
            else:
                room_time.add(kr)

            kp = (p, d, slot)
            if kp in prof_time:
                penalty += PENALTY_PROF_CONFLICT
            else:
                prof_time.add(kp)

            year_time_courses[y][(d, slot)].add(c)

    # حدود الساعات
    for p, used in used_hours.items():
        limit = PROF_LIMITS.get(p, 0.0)
        if used > limit + 1e-9:
            penalty += OVER_LIMIT_HARD * (used - limit)
        score += min(used, limit) * 4

    # فجوات الأساتذة (حسب اليوم فقط)
    prof_day_slots = defaultdict(list)
    for g in chrom:
        slot_val = TIME_MAP.get(g["slot"], None)
        if slot_val is None:
            continue
        for d in g["days"]:
            prof_day_slots[(g["prof"], d)].append(slot_val)
    for key, arr in prof_day_slots.items():
        arr.sort()
        for i in range(1, len(arr)):
            gap_units = arr[i] - arr[i-1] - 1
            if gap_units > 0:
                penalty += gap_units * GAP_WEIGHT

    # تعارض السنة
    for y, grid in year_time_courses.items():
        for (d, slot), cs in grid.items():
            if len(cs) >= 2:
                single_only = [c for c in cs if COURSE_SECTIONS_COUNT.get(c, 0) == 1]
                if len(single_only) >= 2:
                    penalty += PENALTY_YEAR_HARD
                else:
                    penalty += PENALTY_YEAR_SOFT

    # عدالة التوزيع
    if used_hours:
        hours_list = list(used_hours.values())
        avg = sum(hours_list) / len(hours_list)
        var = sum((h - avg) ** 2 for h in hours_list) / len(hours_list)
        penalty += var * VARIANCE_WEIGHT

    # مكافأة عدد الشعب
    score += len(chrom) * ASSIGNED_BONUS

    return score - penalty

# ==========================
# معاملات GA
# ==========================
def tournament_select(pop):
    sample = random.sample(pop, k=TOURNAMENT_K)
    sample.sort(key=lambda ch: fitness(ch), reverse=True)
    return sample[0]

def crossover(parent1, parent2):
    parent1 = [g for g in parent1 if g is not None]
    parent2 = [g for g in parent2 if g is not None]
    n = min(len(parent1), len(parent2))
    if n == 0:
        return parent1[:]
    cut = random.randint(1, n - 1)
    child = parent1[:cut] + parent2[cut:]
    return repair_chromosome(child)

def mutate(chrom):
    if not chrom:
        return chrom
    i = random.randrange(len(chrom))
    ok = try_reassign_gene(chrom, i)
    if not ok and random.random() < MUT_PROB:
        j = random.randrange(len(chrom))
        try_reassign_gene(chrom, j)
    return repair_chromosome(chrom)

# ==========================
# تهيئة البيانات (تُستدعى بعد تحميل الملفات وقبل تشغيل GA)
# ==========================
def rebuild_core_from_inputs():
    global COURSE_GENERAL, COURSE_YEAR, COURSE_KIND
    global SECTIONS, COURSE_SECTIONS_COUNT

    # تنظيف أسماء المواد
    terms_df["المادة"] = terms_df["المادة"].astype(str).str.strip().str.replace('"', '')

    # فلترة term.xlsx حسب الفصل المختار (إن وجد)
    terms_df_filtered = terms_df[terms_df.apply(matches_semester, axis=1)].copy()

    # تخصصات المواد (عام)
    COURSE_GENERAL = {}
    for _, r in terms_df_filtered.iterrows():
        c = str(r["المادة"]).strip()
        gen_list = split_multi(r.get("تخصص عام", ""))
        COURSE_GENERAL[c] = gen_list

    # بناء الشعب
    SECTIONS = []
    COURSE_SECTIONS_COUNT = Counter()
    COURSE_YEAR = {}
    COURSE_KIND = {}

    for _, row in terms_df_filtered.iterrows():
        course = str(row["المادة"]).strip()
        if not course:
            continue
        year = int(row.get("السنة", 0) or 0)
        kind = course_kind(course)

        # تخطَّ المواد التي لا يوجد لها قاعات من نوعها
        rooms_for_kind = ROOMS_BY_KIND.get(kind, [])
        if not rooms_for_kind:
            continue

        COURSE_YEAR[course] = year
        COURSE_KIND[course] = kind

        if DISTRIBUTION_MODE == 1:
            students_total = int(row.get("عدد الطلاب", 0) or 0)
            min_allowed = MIN_LAB if kind == "عملي" else MIN_THEORY
            if students_total < min_allowed:
                continue
            capacities = [ROOM_CAP[r] for r in rooms_for_kind]
            capacities.sort(reverse=True)
            remaining = students_total
            while remaining >= min_allowed:
                for cap in capacities:
                    if remaining < min_allowed:
                        break
                    take = min(cap, remaining)
                    SECTIONS.append({"course": course, "kind": kind, "year": year, "size": take})
                    COURSE_SECTIONS_COUNT[course] += 1
                    remaining -= take
        else:
            num_sections = int(row.get("عدد الشعب", 0) or 0)
            if num_sections <= 0:
                continue
            max_cap = max([ROOM_CAP[r] for r in rooms_for_kind], default=0)
            for _ in range(num_sections):
                SECTIONS.append({"course": course, "kind": kind, "year": year, "size": max_cap})
                COURSE_SECTIONS_COUNT[course] += 1

    if not SECTIONS:
        print("⚠️ لا توجد شعب مطلوبة بعد الفلترة الحالية.")

# ==========================
# تشغيل GA
# ==========================
def run_ga():
    base = random_chromosome()
    base = repair_chromosome(base)

    if not base:
        print("⚠️ لم يتم إسناد أي شعبة في التهيئة. راجع القيود/القاعات/الساعات/فلترة الفصل.")
        return []

    population = [base]
    for _ in range(POP_SIZE - 1):
        population.append(repair_chromosome(mutate(base[:])))

    best = None
    best_fit = float('-inf')
    no_improve = 0

    total_secs = len(SECTIONS)
    print(f"ℹ️ الشعب المطلوبة: {total_secs} | الشعب المسندة في فرد البداية: {len(base)}")

    for gen in range(GENERATIONS):
        population = [[g for g in ch if g is not None] for ch in population]
        population.sort(key=lambda ch: fitness(ch), reverse=True)
        cur_fit = fitness(population[0])

        if cur_fit > best_fit + 1e-6:
            best_fit = cur_fit
            best = population[0]
            no_improve = 0
        else:
            no_improve += 1

        if (gen + 1) % 10 == 0:
            print(f"جيل {gen+1}/{GENERATIONS} | أفضل لياقة: {best_fit:.2f} | مسند: {len(best)} | بدون تحسن: {no_improve}")

        if no_improve >= EARLY_STOP_NO_IMPROVE:
            print(f"⏹️ إيقاف مبكر بعد {gen+1} جيل.")
            break

        next_pop = population[:ELITE]
        while len(next_pop) < POP_SIZE:
            p1 = tournament_select(population)
            p2 = tournament_select(population)
            child = crossover(p1, p2)
            child = mutate(child)
            next_pop.append(child)
        population = next_pop

    return best

# ==========================
# التصدير إلى Excel
# ==========================
def export_schedule(best):
    best = [g for g in best if g is not None]
    used_hours = compute_used_hours(best)

    rows = []
    for g in best:
        prof = g.get("prof")
        total_allowed = PROF_LIMITS.get(prof, 0.0) if prof is not None else 0.0
        used = used_hours.get(prof, 0.0) if prof is not None else 0.0
        days_str = " ".join(g["days"]) if isinstance(g["days"], list) else str(g["days"])
        rows.append({
            "الدكتور": prof,
            "المادة": g["course"],
            "نوع المادة": g["kind"],
            "القاعة": g["room"],
            "عدد الطلاب": g["size"],
            "اليوم": days_str,
            "الوقت": g["slot"],
            "عدد الساعات المستعملة من أصل الكل": f"{used} / {total_allowed}",
            "سنة المادة": g["year"],
        })

    df = pd.DataFrame(rows)
    df = df.sort_values(by=["الدكتور", "اليوم", "الوقت"])
    out = "جدول_محسّن_GA_منع_تعارض_الدكتور.xlsx"
    if os.path.exists(out):
        os.remove(out)
    df.to_excel(out, index=False)
    print(f"✅ تم حفظ الملف: {out}")
    print(f"📦 عدد الشعب المخرجة: {len(df)}")

# ==========================
# التنفيذ
# ==========================

def _bind_exit_on_close(win):
    """أي إغلاق بـ X على هذه النافذة = إنهاء التطبيق كله."""
    def _exit():
        try:
            if APP_ROOT is not None:
                APP_ROOT.destroy()
        except:
            pass
        try:
            win.destroy()
        except:
            pass
        sys.exit(0)
    win.protocol("WM_DELETE_WINDOW", _exit)


# ======== إعدادات/مساعدات RTL وحديثة ========
best_solution = None
file_terms = None
file_profs = None
file_rooms = None
APP_ROOT = None

BASE_FONT_FAMILY = "Cairo"   # جرّب "Cairo" أو "Tahoma" أو "Segoe UI"
BASE_FONT_SIZE = 13

def ar(text: str) -> str:
    """يضمن اتجاه RTL وترتيب علامات الترقيم (مثل :) بشكل صحيح."""
    return "\u202B" + text + "\u202C"

def load_file(var_name):
    path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if path:
        globals()[var_name] = path
        messagebox.showinfo(ar("تم التحميل"), f"{ar('تم تحميل الملف')}: \n{os.path.basename(path)}")

def show_loading_screen(callback):
    # نافذة انتظار صغيرة
    win = tb.Toplevel()
    _bind_exit_on_close(win)  # <<<<<< مهم

    win.title(ar("جاري إنشاء الجدول"))
    win.geometry("350x140")
    win.resizable(False, False)
    win.attributes("-topmost", True)

    container = tb.Frame(win, padding=20)
    container.pack(fill="both", expand=True)

    tb.Label(
        container,
        text=ar("⏳ يرجى الانتظار بينما يتم تجهيز الجدول..."),
        font=(BASE_FONT_FAMILY, BASE_FONT_SIZE+1),
        anchor="e",
        justify="center"
    ).pack(fill="x", pady=(10, 12))

    prog = tb.Progressbar(container, mode="indeterminate", bootstyle="info-striped")
    prog.pack(fill="x")
    prog.start(10)

    def worker():
        err = None
        try:
            callback()
        except Exception as e:
            err = e

        def finish_on_main():
            try: prog.stop()
            except: pass
            try: win.destroy()
            except: pass

            if err is not None:
                try:
                    messagebox.showerror(ar("خطأ أثناء إنشاء الجدول"), str(err))
                except:
                    pass
                # بعد الخطأ، ننهي البرنامج أيضاً
                try:
                    if APP_ROOT is not None:
                        APP_ROOT.destroy()
                except:
                    pass
                sys.exit(0)
            else:
                show_save_screen()

        win.after(0, finish_on_main)

    threading.Thread(target=worker, daemon=True).start()


def show_save_screen():
    win = tb.Toplevel()
    _bind_exit_on_close(win)  # <<<<<< مهم

    win.title(ar("تم تجهيز الجدول"))
    win.geometry("480x200")
    win.resizable(False, False)

    wrap = tb.Frame(win, padding=22, bootstyle="light")
    wrap.pack(fill="both", expand=True)

    tb.Label(
        wrap, text=ar("✅ تم تجهيز الجدول بنجاح"),
        font=(BASE_FONT_FAMILY, BASE_FONT_SIZE+3, "bold"),
        anchor="center", justify="center"
    ).pack(fill="x", pady=(1, 18))

    tb.Button(
        wrap, text=ar("📁 اضغط لحفظ الملف"), bootstyle="success", width=30,
        command=save_file
    ).pack(pady=(0, 8))

def save_file():
    global APP_ROOT
    save_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel file", "*.xlsx")],
        title=ar("أين تريد حفظ الملف؟")
    )
    if save_path:
        try:
            export_schedule(best_solution)
            temp_name = "جدول_محسّن_GA_منع_تعارض_الدكتور.xlsx"
            if os.path.exists(temp_name):
                if os.path.abspath(temp_name) != os.path.abspath(save_path):
                    if os.path.exists(save_path):
                        os.remove(save_path)
                    os.replace(temp_name, save_path)
            messagebox.showinfo(ar("تم الحفظ"), f"{ar('تم حفظ الملف في')}:\n{save_path}")
        except Exception as e:
            messagebox.showerror(ar("خطأ"), str(e))
        finally:
            try:
                if APP_ROOT is not None:
                    APP_ROOT.destroy()  # إغلاق التطبيق نهائيًا بعد الحفظ
            except:
                pass

def start_schedule(college, semester, dist_mode, root):
    global best_solution, SEMESTER_FILTER, DISTRIBUTION_MODE
    if not all([file_terms, file_profs, file_rooms]):
        messagebox.showerror(ar("خطأ"), ar("يرجى تحميل جميع الملفات."))
        return
    if semester not in (1, 2):
        messagebox.showerror(ar("خطأ"), ar("يرجى اختيار فصل دراسي."))
        return

    SEMESTER_FILTER = semester
    DISTRIBUTION_MODE = dist_mode

    # تحميل البيانات من الملفات التي اختارها المستخدم
    global professors_df, rooms_df, terms_df
    professors_df = pd.read_excel(file_profs)
    rooms_df = pd.read_excel(file_rooms)
    terms_df = pd.read_excel(file_terms)

    # إعادة بناء بيانات القاعات/السعات من الملف المحمّل
    global ROOMS_BY_KIND, ROOM_CAP
    ROOMS_BY_KIND = {
        "نظري": rooms_df[rooms_df["النوع"] == "نظري"]["القاعة"].astype(str).tolist(),
        "عملي": rooms_df[rooms_df["النوع"] == "عملي"]["القاعة"].astype(str).tolist(),
    }
    ROOM_CAP = {str(r["القاعة"]): int(r["المساحة"]) for _, r in rooms_df.iterrows()}

    # إعادة بناء بيانات الدكاترة من الملف المحمّل
    global PROF_COURSES, PROF_LIMITS, PROF_GENERAL
    PROF_COURSES = {}
    PROF_LIMITS = {}
    PROF_GENERAL = {}
    for _, r in professors_df.iterrows():
        p = r["الدكتور"]
        PROF_COURSES[p] = set(parse_courses(r.get("المواد", "")))
        try:
            PROF_LIMITS[p] = float(r.get("عدد الساعات", 0) or 0)
        except:
            PROF_LIMITS[p] = 0.0
        PROF_GENERAL[p] = set(split_multi(r.get("التخصصات العامة", "")))

    # إعادة بناء القلب (المواد/الشعب … إلخ) بعد اختيار الفصل وطريقة التوزيع
    rebuild_core_from_inputs()

    # إخفاء النافذة الرئيسية أثناء المعالجة
    root.withdraw()

    def run_ga_thread():
        global best_solution
        best_solution = run_ga()

    show_loading_screen(run_ga_thread)

def launch_gui():
    # ========== إعداد التطبيق والثيم ==========
    global APP_ROOT, BASE_FONT_FAMILY, BASE_FONT_SIZE
    BASE_FONT_FAMILY = "Calibri"

    app = tb.Window(title=ar("نظام الجدولة الأكاديمي"), themename="minty")
    APP_ROOT = app
    _bind_exit_on_close(app)  # <<<<<< مهم

    app.geometry("430x560")
    app.maxsize(430,560)
    app.minsize(430,560)

    style = tb.Style()
    style.configure(".", font=(BASE_FONT_FAMILY, BASE_FONT_SIZE))
    style.configure("TButton", padding=(10, 6), font=(BASE_FONT_FAMILY, BASE_FONT_SIZE, "bold"))
    style.configure("TRadiobutton", padding=(4, 2), font=(BASE_FONT_FAMILY, BASE_FONT_SIZE))
    style.configure("TEntry", padding=6, font=(BASE_FONT_FAMILY, BASE_FONT_SIZE))
    style.configure("Card.TFrame", relief="flat")
    style.configure("Section.TLabel", font=(BASE_FONT_FAMILY, BASE_FONT_SIZE+1, "bold"))
    style.configure("Hint.TLabel", font=(BASE_FONT_FAMILY, BASE_FONT_SIZE-1))
    style.configure("Choice.TRadiobutton", font=(BASE_FONT_FAMILY, BASE_FONT_SIZE, "bold"))

    # ترويسة
    header = tb.Frame(app, padding=(10, 6))
    header.pack(fill="x")
    tb.Label(
        header, text=ar("📅 كلية الهندسة التكنولوجية"),
        anchor="e", justify="right",
        font=(BASE_FONT_FAMILY, BASE_FONT_SIZE+4, "bold")
    ).pack(side="right")
    Separator(app, bootstyle="secondary").pack(fill="x")

    # منطقة المحتوى
    body = tb.Frame(app, padding=12)
    body.pack(fill="both", expand=True)

    # ========== أدوات Hover عامة ==========
    def add_hover_button(btn, normal_bs: str, hover_bs: str):
        btn.configure(bootstyle=normal_bs)
        def _on_enter(_): 
            try: btn.configure(bootstyle=hover_bs)
            except: pass
        def _on_leave(_): 
            try: btn.configure(bootstyle=normal_bs)
            except: pass
        btn.bind("<Enter>", _on_enter)
        btn.bind("<Leave>", _on_leave)
        
    # ---------- بطاقة: المعلومات العامة ----------
    info_card = tb.Frame(body, padding=12, bootstyle="Card")
    info_card.pack(fill="x", pady=3)
    info_card.grid_columnconfigure(0, weight=1)
    info_card.grid_columnconfigure(1, weight=1)

    tb.Label(info_card, text=ar("المعلومات العامة"), style="Section.TLabel",
                anchor="e", justify="right").grid(row=0, column=0, columnspan=2, sticky="e", pady=(0, 4))

    college_var  = tb.StringVar()
    semester_var = tb.IntVar(value=0)
    dist_var     = tb.IntVar(value=2)

    # الفصل الدراسي (Hover)
    tb.Label(info_card, text=ar("الفصل الدراسي:"), anchor="e", justify="right",
                font=(BASE_FONT_FAMILY, BASE_FONT_SIZE, "bold")).grid(
        row=2, column=1, sticky="e", padx=(1, 4), pady=1
    )
    sem_frame = tb.Frame(info_card)
    sem_frame.grid(row=2, column=0, sticky="e", padx=(1, 4), pady=1)

    rb_sem1 = tb.Radiobutton(sem_frame, text=ar("الفصل الأول"), variable=semester_var, value=1)
    rb_sem1.pack(side="right", padx=1)

    rb_sem2 = tb.Radiobutton(sem_frame, text=ar("الفصل الثاني"), variable=semester_var, value=2)
    rb_sem2.pack(side="right", padx=1)

    # طريقة التوزيع (Hover)
    tb.Label(info_card, text=ar("طريقة التوزيع:"), anchor="e", justify="right",
                font=(BASE_FONT_FAMILY, BASE_FONT_SIZE, "bold")).grid(
        row=3, column=1, sticky="e", padx=(1, 4), pady=1
    )
    dist_frame = tb.Frame(info_card)
    dist_frame.grid(row=3, column=0, sticky="e", padx=(1, 4), pady=1)

    rb_dist1 = tb.Radiobutton(dist_frame, text=ar("حسب الطلاب"), variable=dist_var, value=1)
    rb_dist1.pack(side="right", padx=6)

    rb_dist2 = tb.Radiobutton(dist_frame, text=ar("حسب الشعب"), variable=dist_var, value=2)
    rb_dist2.pack(side="right", padx=6)

    # ---------- بطاقة: تحميل الملفات ----------
    files_card = tb.Frame(body, padding=12, bootstyle="Card")
    files_card.pack(fill="x", pady=1)
    files_card.grid_columnconfigure(0, weight=1)

    tb.Label(files_card, text=ar("الملفات المطلوبة"), style="Section.TLabel",
                anchor="e", justify="right").pack(fill="x", pady=(0, 4))
    tb.Label(files_card, text=ar("حمّل الملفات الثلاثة: المواد، الأساتذة، القاعات."),
                style="Hint.TLabel", anchor="e", justify="right").pack(fill="x", pady=(0, 6))

    btn_terms = tb.Button(files_card, text=ar("📄 ملف المواد"),
                            command=lambda: load_file("file_terms"))
    btn_terms.pack(fill="x", pady=1)
    add_hover_button(btn_terms, "secondary-outline", "secondary")

    btn_profs = tb.Button(files_card, text=ar("👨‍🏫 ملف الأساتذة"),
                            command=lambda: load_file("file_profs"))
    btn_profs.pack(fill="x", pady=1)
    add_hover_button(btn_profs, "secondary-outline", "secondary")

    btn_rooms = tb.Button(files_card, text=ar("🏫 ملف القاعات"),
                            command=lambda: load_file("file_rooms"))
    btn_rooms.pack(fill="x", pady=1)
    add_hover_button(btn_rooms, "secondary-outline", "secondary")

    # ---------- بطاقة: التشغيل ----------
    run_card = tb.Frame(body, padding=12, bootstyle="Card")
    run_card.pack(fill="x", pady=1)
    run_card.grid_columnconfigure(0, weight=1)

    tb.Label(run_card, text=ar("جاهز للتشغيل"),
                font=(BASE_FONT_FAMILY, BASE_FONT_SIZE+1, "bold"),
                anchor="center", justify="center").pack(fill="x", pady=(0, 6))

    btn_run = tb.Button(
        run_card, text=ar("🚀 تشغيل الجدولة"),
        command=lambda: start_schedule(college_var.get(), semester_var.get(), dist_var.get(), app)
    )
    btn_run.pack(pady=1, ipady=1)
    add_hover_button(btn_run, "success-outline", "success")

    # تذييل
    footer = tb.Frame(app, padding=8)
    footer.pack(fill="x")
    tb.Label(
        footer, text=ar("نظام الجدولة الأكاديمي ©"),
        font=(BASE_FONT_FAMILY, BASE_FONT_SIZE-1),
        anchor="center", justify="center"
    ).pack(fill="x")

    app.mainloop()

# استدعاء
launch_gui()
