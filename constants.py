"""
펩리치 바이오펩톤 성분 분석 자동화 — 상수 및 매핑 테이블
"""

# 아미노산 25종 (출력 시트 순서 고정)
AMINO_ACIDS = [
    "Aspartic acid",
    "Hydroxyproline",
    "Threonine",
    "Serine",
    "Asparagine",
    "Glutamic acid",
    "Glutamine",
    "Cysteine",
    "Proline",
    "Glycine",
    "Alanine",
    "Citruline",
    "Valine",
    "Cystine",
    "Methionine",
    "Isoleucine",
    "Leucine",
    "Tyrosine",
    "Phenylalanine",
    "GABA",
    "Histidine",
    "Tryptophan",
    "Ornithine",
    "Lysine",
    "Arginine",
]

# 일반 성분 (섹션 1)
GENERAL_COMPONENTS = ["TN", "AN", "총당", "환원당", "회분", "수분", "조지방", "염도"]

# 유리당 (섹션 2)
FREE_SUGARS = ["Fructose", "Glucose", "Sucrose", "Lactose", "Maltose"]

# 미네랄 (섹션 3)
MINERALS = ["Na", "K", "Mg", "Ca"]

# 핵산 (섹션 4)
NUCLEIC_ACIDS = ["AMP", "GMP", "UMP", "IMP", "CMP", "Hypoxantine"]

# 비타민 B (섹션 5)
VITAMIN_B = ["B1", "B2", "B3", "B6", "B9"]

# 성상 항목
SENSANG_ITEMS = ["pH", "탁도", "색도 (L)", "색도 (a)", "색도 (b)"]

# 의뢰품검사상세 검사항목 → 카테고리 매핑
LAB_CATEGORY_MAP = {
    "Vitamin B": "vitB",
    "유기산 (Organic acid)": "organic_acid",
    "유리아미노산": "faa",
    "유리아미노산 (Free amino acid)": "faa",
    "총아미노산": "taa",
    "총아미노산 (Total amino acid)": "taa",
    "미네랄": "mineral",
    "미네랄 (Mineral)": "mineral",
    "유리당": "free_sugar",
    "유리당 (Free sugar)": "free_sugar",
    "핵산": "nucleic_acid",
    "핵산 (Nucleic acid)": "nucleic_acid",
    "일반성분": "general",
    "일반 성분": "general",
}

# 엑셀정리파일 검사항목명 → 표준 성분명 매핑
SUMMARY_TEST_MAP = {
    "TN (총질소)_DUMAS법": "TN",
    "TN (총질소)": "TN",
    "AN (아미노태 질소)": "AN",
    "총당 (Total sugar)": "총당",
    "환원당 (DNS법)": "환원당",
    "NaCl (식염)": "염도",
    "수분 (Moisture)": "수분",
    "회분 (Ash)": "회분",
    "조지방 (Crude Fat)": "조지방",
    "pH": "pH",
    "탁도": "탁도",
    "색도 (L)": "색도 (L)",
    "색도 (a)": "색도 (a)",
    "색도 (b)": "색도 (b)",
}

# 비타민 B 상세 항목명 → 표준명 매핑
VITB_NAME_MAP = {
    "Vitamin B1(Thiamine hydrochloride)": "B1",
    "Vitamin B2(Riboflavin)": "B2",
    "Vitamin B3(Nicotinamide)": "B3",
    "Vitamin B6(Pyridoxin hydrochloride)": "B6",
    "Vitamin B9(Folic acid)": "B9",
}
