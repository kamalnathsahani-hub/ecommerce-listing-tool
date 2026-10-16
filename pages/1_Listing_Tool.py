# ======================================================================
# VC ➜ SC CONVERT ENGINE  (MULTI-SHEET / MULTI-CATEGORY)
# ======================================================================
VC_SKIP_SHEET_WORDS = [
    'instruction', 'guideline', 'definition', 'dropdown', 'valid value', 'example',
    'browse', 'rejection', 'seller info', 'conditional', 'readme', 'read me', 'help',
    'image', 'data def', 'lookup',
]

# (regex, [candidate SC product types in priority order])
SC_PT_RULES = [
    (r'\b(co ?ords?|two piece|2 piece|suit sets?|outfit sets?|suits?)\b', ['SUIT', 'OUTFIT_SET', 'SET']),
    (r'\b(sweat ?shirts?|hoodies?|hoody|hooded|sweaters?|pullovers?)\b', ['SWEATSHIRT', 'SWEATER', 'HOODIE']),
    (r'\b(jeans?|denim|pants?|trousers?|cargos?|joggers?|track ?pants?|chinos?|leggings?|palazzos?|pyjamas?)\b',
     ['PANTS', 'TROUSERS']),
    (r'\bshorts?\b', ['SHORTS']),
    (r'\b(kurtas?|kurtis?|tunics?)\b', ['KURTA', 'TUNIC']),
    (r'\b(dress(es)?|gowns?|frocks?)\b', ['DRESS']),
    (r'\b(jackets?|coats?|blazers?)\b', ['COAT', 'JACKET', 'BLAZER']),
    (r'\bskirts?\b', ['SKIRT']),
    (r'\b(t ?shirts?|tees?|polos?|shirts?|tops?)\b', ['SHIRT', 'TOP']),
]


def match_option(options, candidates):
    """Candidates ko dropdown options se case-insensitive match karta hai."""
    if not options:
        return None
    low = {str(o).strip().lower(): o for o in options}
    norm = {re.sub(r'[\s_\-]+', '', k): v for k, v in low.items()}
    for cand in candidates:
        cl = str(cand).strip().lower()
        if cl in low:
            return low[cl]
        cn = re.sub(r'[\s_\-]+', '', cl)
        if cn in norm:
            return norm[cn]
    for cand in candidates:
        cn = re.sub(r'[\s_\-]+', '', str(cand).lower())
        if len(cn) < 3:
            continue
        for k, v in norm.items():
            if cn in k or k in cn:
                return v
    return None


def infer_sc_product_type(text_sources, options):
    """text_sources me se pehla jo match kare uska SC Product Type return karta hai."""
    for txt in text_sources:
        if not txt:
            continue
        t = str(txt).lower()
        t = re.sub(r'short[\s\-_]*sleeves?', ' ', t)      # "short sleeve" ko shorts na samjhe
        t = re.sub(r'[\-_/|,()#]+', ' ', t)
        for pat, cands in SC_PT_RULES:
            if re.search(pat, t):
                m = match_option(options, cands)
                if m:
                    return m
                if not options:
                    return cands[0]
    return None


def is_vc_listing_sheet_name(name):
    low = name.strip().lower()
    return not any(w in low for w in VC_SKIP_SHEET_WORDS)


def build_vc_sheet_layout(ws_vc, sc_key_to_col, sc_lbl_occ_to_col):
    """Ek VC sheet ka layout. Listing sheet nahi hai to None."""
    lbl_r, key_r, data_r = _detect_amz_sheet_rows(ws_vc, 3, 4, 7)
    vc_to_sc, lbl_counts = {}, {}
    L = {'sku': None, 'model': None, 'asin': None, 'ptype': None, 'itype': None, 'title': None}

    for c in range(1, ws_vc.max_column + 1):
        lbl = str(ws_vc.cell(lbl_r, c).value or '').strip().lower()
        key = str(ws_vc.cell(key_r, c).value or '').strip() if key_r else ''
        nk = _norm_amazon_key(key)

        if L['sku'] is None and (nk.startswith('vendor_sku') or lbl in ('vendor sku', 'sku', 'seller sku')):
            L['sku'] = c
        if L['model'] is None and (nk == 'model_name#1.value' or lbl == 'model name'):
            L['model'] = c
        if L['asin'] is None and (nk == 'merchant_suggested_asin#1.value' or lbl == 'merchant suggested asin'):
            L['asin'] = c
        if L['ptype'] is None and (nk == 'product_type#1.value' or lbl in ('product type', 'feed product type')):
            L['ptype'] = c
        if L['itype'] is None and (nk == 'item_type_name#1.value' or lbl == 'item type name'):
            L['itype'] = c
        if L['title'] is None and (nk == 'item_name#1.value' or lbl == 'item name'):
            L['title'] = c

        nk_target = VC_TO_SC_KEY_ALIASES.get(nk, nk)
        occ = 1
        if lbl:
            occ = lbl_counts.get(lbl, 0) + 1
            lbl_counts[lbl] = occ
        sc_c = None
        if nk_target and nk_target in sc_key_to_col:
            sc_c = sc_key_to_col[nk_target]
        else:
            lbl_target = VC_TO_SC_LABEL_ALIASES.get(lbl, lbl)
            if (lbl_target, occ) in sc_lbl_occ_to_col:
                sc_c = sc_lbl_occ_to_col[(lbl_target, occ)]
        if sc_c:
            vc_to_sc[c] = sc_c

    if L['sku'] is None:
        return None
    return {'data_r': data_r, 'cols': L, 'map': vc_to_sc}


def convert_vc_to_sc(vc_file, sc_file, ctx):
    wb_vc = openpyxl.load_workbook(io.BytesIO(vc_file.getvalue()), data_only=True)

    wb_sc = load_workbook_clean(io.BytesIO(sc_file.getvalue()), sc_file.name)
    ws_sc = None
    for sname in wb_sc.sheetnames:
        if sname.strip().lower() == 'template':
            ws_sc = wb_sc[sname]
            break
    if ws_sc is None:
        ws_sc = pick_template_sheet(wb_sc)
    try:
        ungroup_worksheet(ws_sc)
    except Exception:
        pass

    sc_lbl_r, sc_key_r, sc_data_r = _detect_amz_sheet_rows(ws_sc, 4, 5, 7)

    sc_key_to_col, sc_lbl_occ_to_col, sc_lbl_counts, sc_col_to_lbl = {}, {}, {}, {}
    for c in range(1, ws_sc.max_column + 1):
        lbl = str(ws_sc.cell(sc_lbl_r, c).value or '').strip().lower()
        key = str(ws_sc.cell(sc_key_r, c).value or '').strip() if sc_key_r else ''
        nk = _norm_amazon_key(key)
        sc_col_to_lbl[c] = lbl
        if nk and nk not in sc_key_to_col:
            sc_key_to_col[nk] = c
        if lbl:
            occ = sc_lbl_counts.get(lbl, 0) + 1
            sc_lbl_counts[lbl] = occ
            sc_lbl_occ_to_col[(lbl, occ)] = c

    sc_sku_col = (sc_key_to_col.get('contribution_sku#1.value')
                  or sc_lbl_occ_to_col.get(('sku', 1))
                  or sc_lbl_occ_to_col.get(('seller sku', 1)) or 1)
    sc_prod_type_col = (sc_key_to_col.get('product_type#1.value')
                        or sc_lbl_occ_to_col.get(('product type', 1)))
    sc_action_col = sc_key_to_col.get('::record_action') or sc_lbl_occ_to_col.get(('listing action', 1))
    sc_parentage_col = sc_key_to_col.get('parentage_level#1.value') or sc_lbl_occ_to_col.get(('parentage level', 1))
    sc_parent_sku_col = (sc_key_to_col.get('child_parent_sku_relationship#1.parent_sku')
                         or sc_lbl_occ_to_col.get(('parent sku', 1)))
    sc_var_theme_col = sc_key_to_col.get('variation_theme#1.name') or sc_lbl_occ_to_col.get(('variation theme name', 1))
    sc_title_col = sc_key_to_col.get('item_name#1.value') or sc_lbl_occ_to_col.get(('item name', 1))
    sc_prod_id_type_col = (sc_key_to_col.get('amzn1.volt.ca.product_id_type')
                           or sc_lbl_occ_to_col.get(('product id type', 1)))
    sc_prod_id_val_col = (sc_key_to_col.get('amzn1.volt.ca.product_id_value')
                          or sc_lbl_occ_to_col.get(('product id', 1)))
    sc_model_name_col = sc_key_to_col.get('model_name#1.value') or sc_lbl_occ_to_col.get(('model name', 1))
    sc_condition_col = sc_key_to_col.get('condition_type#1.value') or sc_lbl_occ_to_col.get(('item condition', 1))
    sc_fulfillment_col = (sc_key_to_col.get('fulfillment_availability#1.fulfillment_channel_code')
                          or sc_lbl_occ_to_col.get(('fulfillment channel code (in)', 1)))
    sc_quantity_col = (sc_key_to_col.get('fulfillment_availability#1.quantity')
                       or sc_lbl_occ_to_col.get(('quantity (in)', 1))
                       or sc_lbl_occ_to_col.get(('quantity', 1)))
    sc_shipping_tpl_col = (sc_key_to_col.get('merchant_shipping_group#1.value')
                           or sc_lbl_occ_to_col.get(('shipping template (in)', 1)))

    parent_inherited_sc_cols = {c for c, lbl in sc_col_to_lbl.items() if lbl in PARENT_INHERITED_LABELS}

    # SC template ke Product Type dropdown ke options
    pt_options = get_dropdown_options_for_column(ws_sc, sc_prod_type_col) if sc_prod_type_col else None

    curr_r = sc_data_r
    while curr_r <= ws_sc.max_row and ws_sc.cell(curr_r, sc_sku_col).value not in (None, ''):
        curr_r += 1

    # ---------------- VC ki SAARI listing sheets padho ----------------
    vc_rows_data = []            # (child_sku, model_str, mapped_cells, sheet_name)
    model_first_child_cells = {}
    model_sheet = {}
    seen_child_skus = set()
    sheet_summary = []

    for sname in wb_vc.sheetnames:
        ws_vc = wb_vc[sname]
        if not is_vc_listing_sheet_name(sname):
            sheet_summary.append({'Sheet': sname, 'Status': '⏭️ Skipped (non-listing sheet)',
                                  'Rows': 0, 'Product Type': '', 'Duplicates Skipped': 0})
            continue

        layout = build_vc_sheet_layout(ws_vc, sc_key_to_col, sc_lbl_occ_to_col)
        if not layout:
            sheet_summary.append({'Sheet': sname, 'Status': '⏭️ Skipped (Vendor SKU column nahi mila)',
                                  'Rows': 0, 'Product Type': '', 'Duplicates Skipped': 0})
            continue

        L, vc_to_sc, data_r = layout['cols'], layout['map'], layout['data_r']
        rows_here, dups, pts_used = 0, 0, set()

        for r_vc in range(data_r, ws_vc.max_row + 1):
            sku_val = ws_vc.cell(r_vc, L['sku']).value
            if sku_val in (None, ''):
                continue
            child_sku = str(sku_val).strip()
            if child_sku in seen_child_skus:
                dups += 1
                continue
            seen_child_skus.add(child_sku)

            mapped = {}
            for vc_c, sc_c in vc_to_sc.items():
                v = ws_vc.cell(r_vc, vc_c).value
                if v not in (None, ''):
                    mapped[sc_c] = v
            if L['asin'] and sc_prod_id_val_col:
                av = ws_vc.cell(r_vc, L['asin']).value
                if av not in (None, ''):
                    mapped[sc_prod_id_val_col] = str(av).strip()

            # ---- Product Type decide karo ----
            if sc_prod_type_col:
                vc_pt = ws_vc.cell(r_vc, L['ptype']).value if L['ptype'] else None
                direct = match_option(pt_options, [str(vc_pt)]) if (vc_pt and pt_options) else None
                pt_val = direct or infer_sc_product_type(
                    [sname,
                     ws_vc.cell(r_vc, L['itype']).value if L['itype'] else None,
                     ws_vc.cell(r_vc, L['title']).value if L['title'] else None,
                     vc_pt],
                    pt_options)
                if pt_val:
                    mapped[sc_prod_type_col] = pt_val
                    pts_used.add(str(pt_val))
                else:
                    mapped.pop(sc_prod_type_col, None)

            mv = ws_vc.cell(r_vc, L['model']).value if L['model'] else None
            model_str = str(mv).strip() if mv not in (None, '') else ''
            if model_str and model_str not in model_first_child_cells:
                model_first_child_cells[model_str] = mapped
                model_sheet[model_str] = sname

            vc_rows_data.append((child_sku, model_str, mapped, sname))
            rows_here += 1

        sheet_summary.append({
            'Sheet': sname,
            'Status': '✅ Read' if rows_here else '⚠️ Listing sheet hai par data row nahi mili',
            'Rows': rows_here,
            'Product Type': ", ".join(sorted(pts_used)) or '— (blank)',
            'Duplicates Skipped': dups,
        })

    if not vc_rows_data:
        return None, [], 0, sheet_summary

    # ---------------- Fixed values ----------------
    action_val = "Create or Replace (Full Update)"
    parent_val = pick_dropdown_value(ws_sc, sc_parentage_col, "Parent", "Parent") if sc_parentage_col else "Parent"
    child_val = pick_dropdown_value(ws_sc, sc_parentage_col, "Child", "Child") if sc_parentage_col else "Child"
    var_theme_val = "SIZE/COLOR"
    asin_type_val = "ASIN"
    cond_new_val = "New"
    fulfillment_default_val = "Fulfillment by Merchant (Default)"
    qty_zero_val = 0
    shipping_tpl_val = "Migrated Template"

    parent_position = ctx.get('parent_position', 'before_group')
    seen_models, unique_models_ordered = set(), []
    for _, m_str, _, _ in vc_rows_data:
        if m_str and m_str not in seen_models:
            seen_models.add(m_str)
            unique_models_ordered.append(m_str)

    preview, total_written = [], 0

    def apply_common_fixed_fields(row_idx):
        if sc_action_col:
            ws_sc.cell(row_idx, sc_action_col).value = action_val
        if sc_var_theme_col:
            ws_sc.cell(row_idx, sc_var_theme_col).value = var_theme_val
        if sc_condition_col:
            ws_sc.cell(row_idx, sc_condition_col).value = cond_new_val
        if sc_fulfillment_col:
            ws_sc.cell(row_idx, sc_fulfillment_col).value = fulfillment_default_val
        if sc_quantity_col:
            ws_sc.cell(row_idx, sc_quantity_col).value = qty_zero_val

    def write_parent_row(model_sku):
        nonlocal curr_r, total_written
        ref_cells = model_first_child_cells.get(model_sku) or vc_rows_data[0][2]
        for sc_c in parent_inherited_sc_cols:
            if sc_c in ref_cells and ref_cells[sc_c] not in (None, ''):
                ws_sc.cell(curr_r, sc_c).value = ref_cells[sc_c]
        ws_sc.cell(curr_r, sc_sku_col).value = model_sku
        if sc_prod_type_col and sc_prod_type_col in ref_cells:
            ws_sc.cell(curr_r, sc_prod_type_col).value = ref_cells[sc_prod_type_col]
        if sc_parentage_col:
            ws_sc.cell(curr_r, sc_parentage_col).value = parent_val
        parent_title = ""
        if sc_title_col:
            parent_title = _build_parent_title_from_child(ref_cells.get(sc_title_col, ''), model_sku)
            ws_sc.cell(curr_r, sc_title_col).value = parent_title
        apply_common_fixed_fields(curr_r)
        preview.append({
            'Row': curr_r, 'Source Sheet': model_sheet.get(model_sku, ''), 'SKU': model_sku,
            'Product Type': ref_cells.get(sc_prod_type_col, '') if sc_prod_type_col else '',
            'Listing Action': action_val, 'Parentage Level': parent_val, 'Parent SKU': '',
            'Variation Theme': var_theme_val, 'Product Id Type': '', 'Product Id': '',
            'Shipping Template': '', 'Model Name': '', 'Item Name': parent_title,
        })
        curr_r += 1
        total_written += 1

    def write_child_row(child_sku, model_str, mapped_cells, sheet_name):
        nonlocal curr_r, total_written
        for sc_c, val in mapped_cells.items():
            ws_sc.cell(curr_r, sc_c).value = val
        ws_sc.cell(curr_r, sc_sku_col).value = child_sku
        if sc_parentage_col:
            ws_sc.cell(curr_r, sc_parentage_col).value = child_val
        if sc_parent_sku_col and model_str:
            ws_sc.cell(curr_r, sc_parent_sku_col).value = model_str
        if sc_model_name_col and model_str:
            ws_sc.cell(curr_r, sc_model_name_col).value = model_str
        if sc_prod_id_type_col:
            ws_sc.cell(curr_r, sc_prod_id_type_col).value = asin_type_val
        if sc_shipping_tpl_col:
            ws_sc.cell(curr_r, sc_shipping_tpl_col).value = shipping_tpl_val
        apply_common_fixed_fields(curr_r)
        preview.append({
            'Row': curr_r, 'Source Sheet': sheet_name, 'SKU': child_sku,
            'Product Type': mapped_cells.get(sc_prod_type_col, '') if sc_prod_type_col else '',
            'Listing Action': action_val, 'Parentage Level': child_val, 'Parent SKU': model_str,
            'Variation Theme': var_theme_val, 'Product Id Type': asin_type_val,
            'Product Id': mapped_cells.get(sc_prod_id_val_col, '') if sc_prod_id_val_col else '',
            'Shipping Template': shipping_tpl_val, 'Model Name': model_str,
            'Item Name': mapped_cells.get(sc_title_col, '') if sc_title_col else '',
        })
        curr_r += 1
        total_written += 1

    if parent_position == 'top':
        for m in unique_models_ordered:
            write_parent_row(m)
        for c_sku, m_str, m_cells, shn in vc_rows_data:
            write_child_row(c_sku, m_str, m_cells, shn)
    elif parent_position == 'bottom':
        for c_sku, m_str, m_cells, shn in vc_rows_data:
            write_child_row(c_sku, m_str, m_cells, shn)
        for m in unique_models_ordered:
            write_parent_row(m)
    else:
        inserted = set()
        for c_sku, m_str, m_cells, shn in vc_rows_data:
            if m_str and m_str not in inserted:
                inserted.add(m_str)
                write_parent_row(m_str)
            write_child_row(c_sku, m_str, m_cells, shn)

    output = io.BytesIO()
    wb_sc.save(output)
    output.seek(0)
    return output.getvalue(), preview, total_written, sheet_summary
