with right:
    usage_col, manual_col = st.columns([1.2, 1], gap="large")

    with usage_col:
        st.subheader("사용법")
        st.markdown(
            """
1. **brand** 선택  
2. **item_code / 단품명(국문/영문) / 원산지** 입력  
3. **box_type → box_group 선택 (<span style="color:red;font-weight:700;">템플릿 기준표 참조</span>)**  
4. **실행(개별 입력)** 클릭 → PDF 다운로드  
""",
            unsafe_allow_html=True,
        )

    with manual_col:
        st.subheader("브랜드 매뉴얼")
        st.caption("다운로드해서 포장 규격/박스 타입 확인 후 사용하세요.")

        for b in brand_options:
            manual_path = MANUALS_DIR / f"manual_{b}.pdf"
            brand_ko = BRAND_NAME_KO.get(b, b)

            col1, col2 = st.columns([6, 2], gap="small")
            with col1:
                st.markdown(f"**{brand_ko} 포장박스 매뉴얼**")
            with col2:
                if manual_path.exists():
                    with open(manual_path, "rb") as f:
                        st.download_button(
                            "다운로드",
                            data=f,
                            file_name=manual_path.name,
                            mime="application/pdf",
                            key=f"manual_{b}",
                        )
                else:
                    st.caption("없음")

    # ✅ 템플릿 기준표를 right 영역 안으로 이동
    st.markdown("---")
    st.subheader("템플릿 기준표")
    if TEMPLATE_TABLE_IMG.exists():
        st.image(str(TEMPLATE_TABLE_IMG), use_container_width=True)
    else:
        st.warning("assets/template_table.png 파일이 없습니다.")