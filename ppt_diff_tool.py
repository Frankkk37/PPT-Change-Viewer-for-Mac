#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PPT Diff Tool v0.9 GUI
- Manual A/B .pptx comparison
- Simple local Tkinter UI
- SlideID-first matching + content similarity fallback
- Added/removed/moved/modified slide detection
- Page-number-only change filtering
- v0.9: HTML visual report + short-clause text diff
"""
from __future__ import annotations

import argparse, datetime as _dt, difflib, hashlib, html, json, os, posixpath, re, sys, threading, traceback, zipfile
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple
import xml.etree.ElementTree as ET

NS={
 'p':'http://schemas.openxmlformats.org/presentationml/2006/main',
 'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
 'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
 'rel':'http://schemas.openxmlformats.org/package/2006/relationships',
}

@dataclass
class SlideInfo:
    position:int; slide_id:str; r_id:str; slide_path:str
    text:str; text_chunks:List[str]; text_blocks:List[str]
    text_norm:str; text_hash:str; meaningful_text_hash:str
    image_hashes:List[str]; image_count:int; shape_count:int; graphic_frame_count:int
    semantic_xml_hash:str; content_hash:str

@dataclass
class SlideMatch:
    old_position:int; new_position:int; old_slide_id:str; new_slide_id:str
    method:str; similarity:float; moved:bool

def _sha_b(b:bytes)->str: return hashlib.sha256(b).hexdigest()
def _sha_t(t:str)->str: return hashlib.sha256(t.encode('utf-8','ignore')).hexdigest()
def _clean(t:str)->str: return re.sub(r'\s+',' ',t.replace('\u00a0',' ')).strip()
def _norm(t:str)->str: return _clean(t).lower()
def _digits(t:str)->List[str]: return re.findall(r'\d+',t)
def _mask_digits(t:str)->str: return re.sub(r'\d+','#',t)
def _numeric_like(t:str)->bool: return bool(re.fullmatch(r'[\sPpPage第页/\\|:：\-–—().（）\d]+',t.strip() or 'x'))
def _page_num_like(t:str)->bool:
    t=_clean(t)
    return bool(t and len(t)<=12 and _numeric_like(t) and _digits(t))
def _meaningful(items:List[str])->List[str]: return [x for x in items if not _page_num_like(x)]
def _trunc(t:str,n:int=500)->str:
    t=_clean(t)
    return t if len(t)<=n else t[:n-1]+'…'

def _read_xml(zf:zipfile.ZipFile,path:str)->Optional[ET.Element]:
    try: data=zf.read(path)
    except KeyError: return None
    try: return ET.fromstring(data)
    except ET.ParseError: return None

def _resolve(src_xml:str,target:str)->str:
    target=target.replace('\\','/')
    if target.startswith('/'): return target.lstrip('/')
    return posixpath.normpath(posixpath.join(posixpath.dirname(src_xml),target))

def _rels_path(xml_path:str)->str:
    return posixpath.join(posixpath.dirname(xml_path),'_rels',posixpath.basename(xml_path)+'.rels')

def _read_rels(zf:zipfile.ZipFile,xml_path:str)->Dict[str,Dict[str,str]]:
    root=_read_xml(zf,_rels_path(xml_path))
    if root is None: return {}
    out={}
    for rel in root.findall('rel:Relationship',NS):
        rid=rel.attrib.get('Id'); target=rel.attrib.get('Target')
        if rid and target:
            out[rid]={'target':target,'target_resolved':_resolve(xml_path,target),'type':rel.attrib.get('Type',''),'target_mode':rel.attrib.get('TargetMode','')}
    return out

def _slide_order(zf:zipfile.ZipFile)->List[Tuple[str,str]]:
    root=_read_xml(zf,'ppt/presentation.xml')
    if root is None: raise ValueError('Cannot read ppt/presentation.xml. Is this a valid .pptx file?')
    ans=[]
    for s in root.findall('.//p:sldIdLst/p:sldId',NS):
        sid=s.attrib.get('id',''); rid=s.attrib.get('{%s}id'%NS['r'],'')
        if sid and rid: ans.append((sid,rid))
    return ans

def _text_chunks(root:ET.Element)->List[str]:
    out=[]
    for n in root.findall('.//a:t',NS):
        if n.text is not None:
            c=_clean(n.text)
            if c: out.append(c)
    return out

def _text_blocks(root:ET.Element)->List[str]:
    # First group all text runs inside one DrawingML paragraph.
    # This usually corresponds to a bullet line / textbox paragraph / table-cell paragraph.
    out=[]
    for p in root.findall('.//a:p',NS):
        parts=[]
        for t in p.findall('.//a:t',NS):
            if t.text: parts.append(t.text.replace('\u00a0',' '))
        joined=_clean(''.join(parts))
        if joined: out.append(joined)
    return out or _text_chunks(root)

def _counts(root:ET.Element)->Tuple[int,int,int]:
    return len(root.findall('.//p:sp',NS)),len(root.findall('.//p:pic',NS)),len(root.findall('.//p:graphicFrame',NS))

def _image_hashes(zf:zipfile.ZipFile,slide_path:str)->List[str]:
    out=[]
    for rel in _read_rels(zf,slide_path).values():
        target=rel['target_resolved']; typ=rel.get('type','').lower(); mode=rel.get('target_mode','')
        is_img=('image' in typ) or target.lower().endswith(('.png','.jpg','.jpeg','.gif','.bmp','.tif','.tiff','.emf','.wmf','.svg'))
        if not is_img: continue
        if mode=='External': out.append('external:'+_sha_t(rel.get('target',''))); continue
        try: out.append(_sha_b(zf.read(target)))
        except KeyError: out.append('missing:'+_sha_t(target))
    return sorted(out)

def _semantic_xml_hash(root:ET.Element)->str:
    cloned=ET.fromstring(ET.tostring(root,encoding='utf-8'))
    for n in cloned.findall('.//a:t',NS):
        if n.text is not None and _page_num_like(n.text): n.text='__PAGE_NUMBER__'
    return _sha_b(ET.tostring(cloned,encoding='utf-8'))

def extract_pptx(path:Path)->List[SlideInfo]:
    if not path.exists(): raise FileNotFoundError(path)
    if path.suffix.lower()!='.pptx': raise ValueError(f'Only .pptx is supported: {path}')
    slides=[]
    with zipfile.ZipFile(path,'r') as zf:
        pres_rels=_read_rels(zf,'ppt/presentation.xml')
        for idx,(sid,rid) in enumerate(_slide_order(zf),start=1):
            rel=pres_rels.get(rid)
            if not rel: continue
            sp=rel['target_resolved']; root=_read_xml(zf,sp)
            if root is None: continue
            chunks=_text_chunks(root); blocks=_text_blocks(root)
            text='\n'.join(blocks); text_norm=_norm(text); text_hash=_sha_t(text_norm)
            meaningful_norm=_norm('\n'.join(_meaningful(blocks)))
            meaningful_hash=_sha_t(meaningful_norm)
            shp,img_xml,gf=_counts(root); imgs=_image_hashes(zf,sp); img_count=max(img_xml,len(imgs))
            sem_hash=_semantic_xml_hash(root)
            content_basis=json.dumps({'meaningful_text_norm':meaningful_norm,'image_hashes':imgs,'shape_count':shp,'image_count':img_count,'graphic_frame_count':gf},ensure_ascii=False,sort_keys=True)
            slides.append(SlideInfo(idx,str(sid),rid,sp,text,chunks,blocks,text_norm,text_hash,meaningful_hash,imgs,img_count,shp,gf,sem_hash,_sha_t(content_basis)))
    return slides

def _seq_ratio(a:str,b:str)->float:
    if not a and not b: return 1.0
    if not a or not b: return 0.0
    return difflib.SequenceMatcher(None,a,b).ratio()
def _jacc(a:List[str],b:List[str])->float:
    A=set(a); B=set(b)
    if not A and not B: return 1.0
    if not A or not B: return 0.0
    return len(A&B)/len(A|B)
def _cntsim(a:int,b:int)->float:
    if a==b: return 1.0
    hi=max(a,b)
    return 1.0 if hi==0 else 1.0-abs(a-b)/hi

def slide_similarity(o:SlideInfo,n:SlideInfo)->float:
    ot=_norm('\n'.join(_meaningful(o.text_blocks))); nt=_norm('\n'.join(_meaningful(n.text_blocks)))
    text_sim=_seq_ratio(ot,nt); image_sim=_jacc(o.image_hashes,n.image_hashes)
    count_sim=_cntsim(o.shape_count,n.shape_count)*0.45+_cntsim(o.image_count,n.image_count)*0.35+_cntsim(o.graphic_frame_count,n.graphic_frame_count)*0.20
    if len(ot)>=20 or len(nt)>=20: score=0.75*text_sim+0.15*image_sim+0.10*count_sim
    else: score=0.35*text_sim+0.45*image_sim+0.20*count_sim
    return round(score,6)

def match_slides(old:List[SlideInfo],new:List[SlideInfo],threshold:float=0.78):
    old_id={s.slide_id:s for s in old}; new_id={s.slide_id:s for s in new}
    matched_o=set(); matched_n=set(); matches=[]
    for sid,o in old_id.items():
        n=new_id.get(sid)
        if not n: continue
        matched_o.add(o.slide_id); matched_n.add(n.slide_id)
        sim=1.0 if o.content_hash==n.content_hash else slide_similarity(o,n)
        matches.append(SlideMatch(o.position,n.position,o.slide_id,n.slide_id,'slide_id',sim,o.position!=n.position))
    uo=[s for s in old if s.slide_id not in matched_o]; un=[s for s in new if s.slide_id not in matched_n]
    cand=[]
    for o in uo:
        for n in un:
            sim=slide_similarity(o,n)
            if sim>=threshold: cand.append((sim,o,n))
    cand.sort(key=lambda x:x[0],reverse=True)
    used_o=set(); used_n=set()
    for sim,o,n in cand:
        if o.slide_id in used_o or n.slide_id in used_n: continue
        used_o.add(o.slide_id); used_n.add(n.slide_id)
        matches.append(SlideMatch(o.position,n.position,o.slide_id,n.slide_id,'content_similarity',sim,o.position!=n.position))
    removed=[s for s in uo if s.slide_id not in used_o]
    added=[s for s in un if s.slide_id not in used_n]
    matches.sort(key=lambda m:(m.new_position,m.old_position)); added.sort(key=lambda s:s.position); removed.sort(key=lambda s:s.position)
    return matches,removed,added

def _char_ops(old:str,new:str,max_ops:int=30)->List[Dict]:
    ops=[]
    for tag,i1,i2,j1,j2 in difflib.SequenceMatcher(None,old,new).get_opcodes():
        if tag=='equal': continue
        ops.append({'type':tag,'old':old[i1:i2],'new':new[j1:j2],'old_span':[i1,i2],'new_span':[j1,j2]})
        if len(ops)>=max_ops: break
    return ops

def _kind(ops:List[Dict])->str:
    if not ops: return 'text_replaced'
    ins=any(o['type']=='insert' for o in ops); dele=any(o['type']=='delete' for o in ops); repl=any(o['type']=='replace' for o in ops)
    if ins and not dele and not repl: return 'text_added_within_sentence'
    if dele and not ins and not repl: return 'text_removed_within_sentence'
    if repl and not ins and not dele: return 'text_replaced_within_sentence'
    return 'text_changed_within_sentence'

def _split_short_clauses(text:str)->List[str]:
    """
    Split a paragraph/bullet line into short human-readable clauses.

    Boundary punctuation is retained in the segment. This keeps context like
    Chinese commas/semicolons while avoiding very long full-sentence diffs.
    """
    text=_clean(text)
    if not text: return []
    boundaries=set('，,、；;：:。.!！?？')
    # Treat these as soft boundaries too. Keep the char at the end of segment.
    soft=set('/|')
    out=[]; buf=[]
    for ch in text:
        buf.append(ch)
        if ch in boundaries or ch in soft:
            seg=_clean(''.join(buf))
            if seg: out.append(seg)
            buf=[]
    tail=_clean(''.join(buf))
    if tail: out.append(tail)

    # If a segment is still too long, keep it but later render local context.
    return out or [text]

def _context_around_ops(old:str,new:str,ops:List[Dict],radius:int=34)->Tuple[str,str]:
    """
    For unusually long clauses, return shorter context around the first changed span.
    """
    if not ops:
        return _trunc(old,90), _trunc(new,90)
    old_positions=[]; new_positions=[]
    for op in ops:
        os=op.get('old_span') or [0,0]; ns=op.get('new_span') or [0,0]
        old_positions.extend(os); new_positions.extend(ns)
    o_mid=sum(old_positions)//len(old_positions) if old_positions else 0
    n_mid=sum(new_positions)//len(new_positions) if new_positions else 0
    def cut(t,mid):
        if len(t)<=96: return t
        a=max(0,mid-radius); b=min(len(t),mid+radius)
        return ('…' if a>0 else '') + t[a:b] + ('…' if b<len(t) else '')
    return cut(old,o_mid), cut(new,n_mid)

def _segment_changes_for_pair(old_txt:str,new_txt:str,base_old_idx:int,base_new_idx:int,max_changes:int)->List[Dict]:
    """
    Compare two paragraph blocks at short-clause level.
    """
    old_segs=_split_short_clauses(old_txt)
    new_segs=_split_short_clauses(new_txt)
    sm=difflib.SequenceMatcher(None,old_segs,new_segs)
    out=[]
    def add(c):
        if len(out)<max_changes: out.append(c)
    for tag,i1,i2,j1,j2 in sm.get_opcodes():
        if tag=='equal': continue
        os=old_segs[i1:i2]; ns=new_segs[j1:j2]
        if tag=='delete':
            for k,seg in enumerate(os,start=i1+1):
                add({'type':'removed_segment','old_block_index':base_old_idx,'new_block_index':base_new_idx,'old_segment_index':k,'new_segment_index':None,'old':_trunc(seg,180),'new':'','char_changes':[],'change_kind':'removed_short_clause'})
        elif tag=='insert':
            for k,seg in enumerate(ns,start=j1+1):
                add({'type':'added_segment','old_block_index':base_old_idx,'new_block_index':base_new_idx,'old_segment_index':None,'new_segment_index':k,'old':'','new':_trunc(seg,180),'char_changes':[],'change_kind':'added_short_clause'})
        elif tag=='replace':
            for k in range(max(len(os),len(ns))):
                old_seg=os[k] if k<len(os) else ''
                new_seg=ns[k] if k<len(ns) else ''
                if old_seg and new_seg:
                    ops=_char_ops(old_seg,new_seg)
                    old_show,new_show=_context_around_ops(old_seg,new_seg,ops) if (len(old_seg)>96 or len(new_seg)>96) else (old_seg,new_seg)
                    add({'type':'changed_segment','old_block_index':base_old_idx,'new_block_index':base_new_idx,'old_segment_index':i1+k+1,'new_segment_index':j1+k+1,'old':_trunc(old_show,180),'new':_trunc(new_show,180),'char_changes':ops,'change_kind':_kind(ops),'full_old':_trunc(old_seg,420),'full_new':_trunc(new_seg,420)})
                elif old_seg:
                    add({'type':'removed_segment','old_block_index':base_old_idx,'new_block_index':base_new_idx,'old_segment_index':i1+k+1,'new_segment_index':None,'old':_trunc(old_seg,180),'new':'','char_changes':[],'change_kind':'removed_short_clause'})
                else:
                    add({'type':'added_segment','old_block_index':base_old_idx,'new_block_index':base_new_idx,'old_segment_index':None,'new_segment_index':j1+k+1,'old':'','new':_trunc(new_seg,180),'char_changes':[],'change_kind':'added_short_clause'})
    return out

def build_text_changes(o:SlideInfo,n:SlideInfo,max_changes:int=30)->List[Dict]:
    """
    v0.9 human-readable text diff.

    Pipeline:
    1. Merge PowerPoint XML runs into paragraph/bullet blocks.
    2. Match changed paragraph blocks.
    3. Split each changed paragraph into short clauses by punctuation.
    4. Show only changed short clauses, with char-level change summary.

    This avoids both extremes:
    - raw XML text-run fragments that are too short;
    - full paragraph/sentence diffs that are too long.
    """
    old_blocks=_meaningful(o.text_blocks) or _meaningful(o.text_chunks)
    new_blocks=_meaningful(n.text_blocks) or _meaningful(n.text_chunks)
    sm=difflib.SequenceMatcher(None,old_blocks,new_blocks)
    changes=[]
    def add(c):
        if len(changes)<max_changes: changes.append(c)
    for tag,i1,i2,j1,j2 in sm.get_opcodes():
        if tag=='equal': continue
        ob=old_blocks[i1:i2]; nb=new_blocks[j1:j2]
        if tag=='delete':
            for idx,txt in enumerate(ob,start=i1+1):
                for k,seg in enumerate(_split_short_clauses(txt),start=1):
                    add({'type':'removed_segment','old_block_index':idx,'new_block_index':None,'old_segment_index':k,'new_segment_index':None,'old':_trunc(seg,180),'new':'','char_changes':[],'change_kind':'removed_short_clause'})
        elif tag=='insert':
            for idx,txt in enumerate(nb,start=j1+1):
                for k,seg in enumerate(_split_short_clauses(txt),start=1):
                    add({'type':'added_segment','old_block_index':None,'new_block_index':idx,'old_segment_index':None,'new_segment_index':k,'old':'','new':_trunc(seg,180),'char_changes':[],'change_kind':'added_short_clause'})
        elif tag=='replace':
            for k in range(max(len(ob),len(nb))):
                old_txt=ob[k] if k<len(ob) else ''
                new_txt=nb[k] if k<len(nb) else ''
                if old_txt and new_txt:
                    seg_changes=_segment_changes_for_pair(old_txt,new_txt,i1+k+1,j1+k+1,max_changes-len(changes))
                    if seg_changes:
                        for c in seg_changes: add(c)
                    else:
                        ops=_char_ops(old_txt,new_txt)
                        add({'type':'changed_segment','old_block_index':i1+k+1,'new_block_index':j1+k+1,'old_segment_index':None,'new_segment_index':None,'old':_trunc(old_txt,180),'new':_trunc(new_txt,180),'char_changes':ops,'change_kind':_kind(ops),'full_old':_trunc(old_txt,420),'full_new':_trunc(new_txt,420)})
                elif old_txt:
                    for si,seg in enumerate(_split_short_clauses(old_txt),start=1):
                        add({'type':'removed_segment','old_block_index':i1+k+1,'new_block_index':None,'old_segment_index':si,'new_segment_index':None,'old':_trunc(seg,180),'new':'','char_changes':[],'change_kind':'removed_short_clause'})
                else:
                    for si,seg in enumerate(_split_short_clauses(new_txt),start=1):
                        add({'type':'added_segment','old_block_index':None,'new_block_index':j1+k+1,'old_segment_index':None,'new_segment_index':si,'old':'','new':_trunc(seg,180),'char_changes':[],'change_kind':'added_short_clause'})
    if len(changes)>=max_changes:
        changes.append({'type':'truncated','old':'','new':f'Text diff truncated at {max_changes} changes.','char_changes':[],'change_kind':'truncated'})
    return changes

def _page_num_change(c:Dict)->bool:
    old=(c.get('old') or '').strip(); new=(c.get('new') or '').strip()
    if c.get('type')=='truncated': return False
    if old and new:
        if old==new: return True
        if _digits(old) and _digits(new) and _mask_digits(old)==_mask_digits(new): return True
        if _page_num_like(old) and _page_num_like(new): return True
        return False
    return _page_num_like(old or new)

def build_moved_groups(items:List[Dict])->List[Dict]:
    if not items: return []
    items=sorted(items,key=lambda x:(x['old_position'],x['new_position']))
    groups=[]; cur=[items[0]]
    def same(a,b): return b['old_position']==a['old_position']+1 and b['new_position']==a['new_position']+1 and (b['new_position']-b['old_position'])==(a['new_position']-a['old_position'])
    for it in items[1:]:
        if same(cur[-1],it): cur.append(it)
        else: groups.append(cur); cur=[it]
    groups.append(cur)
    out=[]
    for g in groups:
        delta=g[0]['new_position']-g[0]['old_position']
        out.append({'old_start':g[0]['old_position'],'old_end':g[-1]['old_position'],'new_start':g[0]['new_position'],'new_end':g[-1]['new_position'],'count':len(g),'delta':delta,'direction':'后移' if delta>0 else '前移' if delta<0 else '未移动','items':g})
    out.sort(key=lambda x:(-x['count'],x['old_start'],x['new_start']))
    return out

def build_diff(old_path:Path,new_path:Path,detect_format:bool=False,threshold:float=0.78)->Dict:
    old=extract_pptx(old_path); new=extract_pptx(new_path); matches,removed,added=match_slides(old,new,threshold)
    old_id={s.slide_id:s for s in old}; new_id={s.slide_id:s for s in new}; old_pos={s.position:s for s in old}; new_pos={s.position:s for s in new}
    moved=[]; modified=[]; ignored=[]; format_only=[]
    for m in matches:
        o=old_id.get(m.old_slide_id) or old_pos.get(m.old_position); n=new_id.get(m.new_slide_id) or new_pos.get(m.new_position)
        if not o or not n: continue
        if m.moved:
            moved.append({'old_position':m.old_position,'new_position':m.new_position,'old_slide_id':m.old_slide_id,'new_slide_id':m.new_slide_id,'match_method':m.method,'similarity':m.similarity,'delta':m.new_position-m.old_position})
        raw_text=o.text_hash!=n.text_hash; meaningful=o.meaningful_text_hash!=n.meaningful_text_hash
        imgs=o.image_hashes!=n.image_hashes; shp=o.shape_count!=n.shape_count; imgc=o.image_count!=n.image_count; gfc=o.graphic_frame_count!=n.graphic_frame_count
        sem=o.semantic_xml_hash!=n.semantic_xml_hash
        text_changes=build_text_changes(o,n) if raw_text else []
        only_page=raw_text and text_changes and all(_page_num_change(c) for c in text_changes)
        content=any([meaningful,imgs,shp,imgc,gfc])
        fmt=bool(detect_format and sem and not content and not only_page)
        detail={'old_position':m.old_position,'new_position':m.new_position,'old_slide_id':m.old_slide_id,'new_slide_id':m.new_slide_id,'match_method':m.method,'similarity':m.similarity,
                'changed':{'meaningful_text':meaningful,'raw_text':raw_text,'images':imgs,'shape_count':shp,'image_count':imgc,'graphic_frame_count':gfc,'format_or_layout_xml':fmt,'only_page_number_text_change':only_page},
                'old_counts':{'characters':len(o.text_norm),'meaningful_characters':len(_norm('\n'.join(_meaningful(o.text_blocks)))),'text_chunks':len(o.text_chunks),'text_blocks':len(o.text_blocks),'images':o.image_count,'shapes':o.shape_count,'graphic_frames':o.graphic_frame_count},
                'new_counts':{'characters':len(n.text_norm),'meaningful_characters':len(_norm('\n'.join(_meaningful(n.text_blocks)))),'text_chunks':len(n.text_chunks),'text_blocks':len(n.text_blocks),'images':n.image_count,'shapes':n.shape_count,'graphic_frames':n.graphic_frame_count},
                'text_changes':text_changes if meaningful else []}
        if content or fmt:
            modified.append(detail)
            if fmt and not content: format_only.append(detail)
        elif only_page: ignored.append(detail)
    groups=build_moved_groups(moved)
    return {'schema_version':'0.9','generated_at':_dt.datetime.now().isoformat(timespec='seconds'),'from_file':str(old_path),'to_file':str(new_path),'from_name':old_path.stem,'to_name':new_path.stem,
            'settings':{'similarity_threshold':threshold,'matching':'slide_id_first_then_content_similarity','page_number_filter':True,'detect_format_or_layout':detect_format,'text_diff_mode':'short_clause'},
            'summary':{'old_slide_count':len(old),'new_slide_count':len(new),'added':len(added),'removed':len(removed),'moved':len(moved),'moved_groups':len(groups),'modified':len(modified),'format_or_layout_only':len(format_only),'matched':len(matches),'ignored_page_number_only_changes':len(ignored)},
            'added_slides':[{'new_position':s.position,'new_slide_id':s.slide_id,'text_preview':_trunc(s.text_norm,160),'image_count':s.image_count,'shape_count':s.shape_count} for s in added],
            'removed_slides':[{'old_position':s.position,'old_slide_id':s.slide_id,'text_preview':_trunc(s.text_norm,160),'image_count':s.image_count,'shape_count':s.shape_count} for s in removed],
            'moved_slides':moved,'moved_groups':groups,'modified_slides':modified,'format_only_slides':format_only,'ignored_page_number_only_changes':ignored,'matches':[asdict(x) for x in matches]}

def safe_filename(name:str)->str: return (re.sub(r'[^\w\-.一-龥]+','_',name,flags=re.UNICODE).strip('_') or 'ppt_diff')
def _range(prefix,s,e): return f'{prefix}{s}' if s==e else f'{prefix}{s}–{prefix}{e}'
def _fmt_group(g):
    delta=abs(g['delta']); old=_range('P',g['old_start'],g['old_end']); new=_range('P',g['new_start'],g['new_end'])
    return f'旧 {old} → 新 {new}（{g["direction"]} {delta} 页）' if g['count']==1 else f'旧 {old} → 新 {new}（整体{g["direction"]} {delta} 页，{g["count"]} 页）'
def _fmt_ops(ops):
    out=[]
    for c in ops[:12]:
        typ=c.get('type'); old=c.get('old',''); new=c.get('new','')
        if typ=='replace': out.append(f'`{old}` → `{new}`')
        elif typ=='delete': out.append(f'删除 `{old}`')
        elif typ=='insert': out.append(f'新增 `{new}`')
        else: out.append(f'`{old}` → `{new}`')
    return '；'.join(out)
def _kind_cn(k):
    return {'text_added_within_sentence':'句内新增','text_removed_within_sentence':'句内删除','text_replaced_within_sentence':'句内替换','text_changed_within_sentence':'句内修改','text_replaced':'文本替换','added_sentence_or_block':'新增句/段','removed_sentence_or_block':'删除句/段','added_short_clause':'新增短句','removed_short_clause':'删除短句'}.get(k,k)

def render_markdown(diff:Dict)->str:
    s=diff['summary']; lines=[]
    lines += ['# PPT Diff Report','',f"- From: `{Path(diff['from_file']).name}`",f"- To: `{Path(diff['to_file']).name}`",f"- Generated at: `{diff['generated_at']}`",f"- Schema version: `{diff['schema_version']}`",f"- Text diff mode: `{diff['settings']['text_diff_mode']}`",f"- Detect format/layout: `{diff['settings']['detect_format_or_layout']}`",'','## Summary','']
    for key,label in [('old_slide_count','Old slide count'),('new_slide_count','New slide count'),('added','Added'),('removed','Removed'),('moved','Moved slides'),('moved_groups','Moved groups'),('modified','Modified'),('format_or_layout_only','Format/layout-only changes'),('ignored_page_number_only_changes','Ignored page-number-only changes')]: lines.append(f'- {label}: **{s[key]}**')
    lines += ['','## Added Slides','']
    lines += [f"- `+ 新 P{x['new_position']}` ｜ {x['text_preview'] or '(no text)'}" for x in diff['added_slides']] or ['- None']
    lines += ['','## Removed Slides','']
    lines += [f"- `- 旧 P{x['old_position']}` ｜ {x['text_preview'] or '(no text)'}" for x in diff['removed_slides']] or ['- None']
    lines += ['','## Moved Slides','']
    if diff['moved_groups']:
        lines += ['### Readable movement groups','']
        lines += [f'- {_fmt_group(g)}' for g in diff['moved_groups']]
        items=sorted(diff['moved_slides'],key=lambda x:(x['old_position'],x['new_position']))
        if len(items)<=30:
            lines += ['','### Full moved-slide details','']
            for it in items:
                d=it['delta']; direction='后移' if d>0 else '前移' if d<0 else '未移动'
                lines.append(f"- `旧 P{it['old_position']} → 新 P{it['new_position']}` ｜ {direction} {abs(d)} 页")
        else: lines.append('- Full moved-slide details are stored in JSON.')
    else: lines.append('- None')
    lines += ['','## Modified Slides','']
    if diff['modified_slides']:
        for item in diff['modified_slides']:
            ch=item['changed']; parts=[]
            if ch.get('meaningful_text'): parts.append('文字')
            if ch.get('images'): parts.append('图片')
            if ch.get('shape_count'): parts.append('形状数量')
            if ch.get('image_count'): parts.append('图片数量')
            if ch.get('graphic_frame_count'): parts.append('表格/图表容器数量')
            if ch.get('format_or_layout_xml'): parts.append('格式/布局')
            lines += [f"### 旧 P{item['old_position']} → 新 P{item['new_position']}",'',f"- Changed: **{'、'.join(parts) if parts else '内容变化'}**",f"- Match method: `{item['match_method']}` ｜ similarity `{item['similarity']}`"]
            oc=item.get('old_counts',{}); nc=item.get('new_counts',{})
            if oc and nc: lines.append(f"- Counts: text blocks {oc.get('text_blocks')} → {nc.get('text_blocks')}, text chunks {oc.get('text_chunks')} → {nc.get('text_chunks')}, text chars {oc.get('characters')} → {nc.get('characters')}, images {oc.get('images')} → {nc.get('images')}, shapes {oc.get('shapes')} → {nc.get('shapes')}, graphic frames {oc.get('graphic_frames')} → {nc.get('graphic_frames')}")
            tc=item.get('text_changes',[])
            if tc:
                lines += ['','Text changes:']
                for c in tc:
                    if c.get('type')=='truncated': lines.append(f"- {c.get('new')}"); continue
                    old=c.get('old',''); new=c.get('new',''); ops=_fmt_ops(c.get('char_changes',[])); kind=_kind_cn(c.get('change_kind',''))
                    if c.get('type')=='changed_segment':
                        lines += [f'- {kind}：',f'  - 旧：{old}',f'  - 新：{new}']
                        if ops: lines.append(f'  - 变化：{ops}')
                        if c.get('full_old') and (c.get('full_old')!=old or c.get('full_new')!=new):
                            lines.append(f'  - 完整短句：{c.get("full_old")} → {c.get("full_new")}')
                    elif c.get('type')=='removed_segment': lines.append(f'- 删除短句：{old}')
                    elif c.get('type')=='added_segment': lines.append(f'- 新增短句：{new}')
                    elif c.get('type')=='changed_block':
                        lines += [f'- {kind}：',f'  - 旧句：{old}',f'  - 新句：{new}']
                        if ops: lines.append(f'  - 变化：{ops}')
                    elif c.get('type')=='removed_block': lines.append(f'- 删除句/段：{old}')
                    elif c.get('type')=='added_block': lines.append(f'- 新增句/段：{new}')
            lines.append('')
    else: lines.append('- None')
    lines += ['','## Notes','','- v0.9 outputs Markdown, JSON, and an HTML visual report. Text changes still use short-clause diff.','- This should make changes easier to spot without reading a full long sentence.','- Page-number-only changes are ignored to avoid false positives after insert/delete/move.','- Format/layout detection is available but disabled by default in v0.9, because it is more sensitive.','']
    return '\n'.join(lines)


def _h(x)->str:
    return html.escape(str(x), quote=True)

def _highlight_html(text:str, ops:List[Dict], side:str)->str:
    if not text:
        return ''
    spans=[]
    key='old_span' if side=='old' else 'new_span'
    cls='del' if side=='old' else 'ins'
    for op in ops or []:
        a,b=op.get(key,[0,0])
        if a==b:
            continue
        a=max(0,min(len(text),a)); b=max(0,min(len(text),b))
        if a<b:
            spans.append((a,b,cls))
    if not spans:
        return _h(text)
    spans.sort()
    merged=[]
    for a,b,c in spans:
        if merged and a<=merged[-1][1]:
            merged[-1]=(merged[-1][0],max(merged[-1][1],b),c)
        else:
            merged.append((a,b,c))
    out=[]; pos=0
    for a,b,c in merged:
        out.append(_h(text[pos:a])); out.append(f'<mark class="{c}">{_h(text[a:b])}</mark>'); pos=b
    out.append(_h(text[pos:]))
    return ''.join(out)

def _changed_labels(item:Dict)->str:
    ch=item.get('changed',{}); parts=[]
    if ch.get('meaningful_text'): parts.append('文字')
    if ch.get('images'): parts.append('图片')
    if ch.get('shape_count'): parts.append('形状数量')
    if ch.get('image_count'): parts.append('图片数量')
    if ch.get('graphic_frame_count'): parts.append('表格/图表容器数量')
    if ch.get('format_or_layout_xml'): parts.append('格式/布局')
    return '、'.join(parts) if parts else '内容变化'

def _render_text_change_html(c:Dict)->str:
    ctype=c.get('type'); kind=_kind_cn(c.get('change_kind',''))
    old=c.get('old',''); new=c.get('new',''); ops=c.get('char_changes',[])
    if ctype=='truncated':
        return f'<div class="change muted">{_h(c.get("new",""))}</div>'
    if ctype=='changed_segment':
        old_html=_highlight_html(old,ops,'old'); new_html=_highlight_html(new,ops,'new')
        full=''
        if c.get('full_old') and (c.get('full_old')!=old or c.get('full_new')!=new):
            full=f'<div class="full">完整短句：{_h(c.get("full_old",""))} → {_h(c.get("full_new",""))}</div>'
        return f'''<div class="change">
  <div class="kind">{_h(kind)}</div>
  <div class="row"><div class="label old">旧</div><div class="text oldbox">{old_html}</div></div>
  <div class="row"><div class="label new">新</div><div class="text newbox">{new_html}</div></div>
  {full}
</div>'''
    if ctype=='added_segment':
        return f'<div class="change"><div class="kind">新增短句</div><div class="row"><div class="label new">新</div><div class="text newbox"><mark class="ins">{_h(new)}</mark></div></div></div>'
    if ctype=='removed_segment':
        return f'<div class="change"><div class="kind">删除短句</div><div class="row"><div class="label old">旧</div><div class="text oldbox"><mark class="del">{_h(old)}</mark></div></div></div>'
    if ctype=='changed_block':
        return f'''<div class="change">
  <div class="kind">{_h(kind)}</div>
  <div class="row"><div class="label old">旧句</div><div class="text oldbox">{_highlight_html(old,ops,'old')}</div></div>
  <div class="row"><div class="label new">新句</div><div class="text newbox">{_highlight_html(new,ops,'new')}</div></div>
</div>'''
    if ctype=='added_block':
        return f'<div class="change"><div class="kind">新增句/段</div><div class="text newbox"><mark class="ins">{_h(new)}</mark></div></div>'
    if ctype=='removed_block':
        return f'<div class="change"><div class="kind">删除句/段</div><div class="text oldbox"><mark class="del">{_h(old)}</mark></div></div>'
    return f'<div class="change"><div>{_h(ctype)}: {_h(old)} → {_h(new)}</div></div>'

def render_html(diff:Dict)->str:
    s=diff['summary']
    title=f"PPT Diff Report — {Path(diff['from_file']).name} → {Path(diff['to_file']).name}"
    def card(label,key,cls=''):
        return f'<div class="card {cls}"><div class="num">{_h(s.get(key,0))}</div><div class="cap">{_h(label)}</div></div>'
    added=''.join([f'<li><b>新 P{x["new_position"]}</b><span>{_h(x.get("text_preview") or "(no text)")}</span></li>' for x in diff.get('added_slides',[])]) or '<li class="muted">None</li>'
    removed=''.join([f'<li><b>旧 P{x["old_position"]}</b><span>{_h(x.get("text_preview") or "(no text)")}</span></li>' for x in diff.get('removed_slides',[])]) or '<li class="muted">None</li>'
    moved=''.join([f'<li>{_h(_fmt_group(g))}</li>' for g in diff.get('moved_groups',[])]) or '<li class="muted">None</li>'
    mods=[]
    for item in diff.get('modified_slides',[]):
        tc=''.join(_render_text_change_html(c) for c in item.get('text_changes',[])) or '<div class="muted">No text-level details.</div>'
        mods.append(f'''<details class="slide" open>
  <summary><span class="pill">旧 P{item['old_position']} → 新 P{item['new_position']}</span><span>{_h(_changed_labels(item))}</span><span class="sim">sim {_h(item.get('similarity'))}</span></summary>
  <div class="slide-body">{tc}</div>
</details>''')
    mods_html=''.join(mods) or '<div class="empty">None</div>'
    return f'''<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{_h(title)}</title>
<style>
:root {{ --bg:#f7f8fb; --panel:#fff; --text:#172033; --muted:#667085; --border:#d8dee9; --blue:#1f6feb; --green:#177245; --red:#b42318; --amber:#b54708; }}
* {{ box-sizing:border-box; }}
body {{ margin:0; background:var(--bg); color:var(--text); font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","PingFang SC","Microsoft YaHei",Arial,sans-serif; line-height:1.55; }}
header {{ padding:28px 34px 18px; background:linear-gradient(135deg,#152238,#263b63); color:white; }}
h1 {{ margin:0 0 8px; font-size:24px; }}
.meta {{ color:#d0d7e2; font-size:13px; }}
main {{ max-width:1120px; margin:0 auto; padding:24px; }}
.cards {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(135px,1fr)); gap:12px; margin-bottom:22px; }}
.card {{ background:var(--panel); border:1px solid var(--border); border-radius:14px; padding:14px 16px; box-shadow:0 1px 2px rgba(16,24,40,.04); }}
.num {{ font-size:26px; font-weight:750; }} .cap {{ color:var(--muted); font-size:13px; }}
.card.add .num {{ color:var(--green); }} .card.rem .num {{ color:var(--red); }} .card.mov .num {{ color:var(--blue); }} .card.mod .num {{ color:var(--amber); }}
section {{ background:var(--panel); border:1px solid var(--border); border-radius:14px; padding:18px 20px; margin:16px 0; }}
h2 {{ margin:0 0 12px; font-size:18px; }}
ul {{ margin:0; padding-left:20px; }} li {{ margin:6px 0; }} li span {{ margin-left:10px; color:var(--muted); }}
.slide {{ border:1px solid var(--border); border-radius:12px; margin:12px 0; background:#fff; overflow:hidden; }}
.slide summary {{ cursor:pointer; padding:13px 15px; display:flex; gap:12px; align-items:center; border-bottom:1px solid var(--border); }}
.pill {{ background:#eef4ff; color:#1849a9; border-radius:999px; padding:3px 9px; font-weight:650; }}
.sim {{ margin-left:auto; color:var(--muted); font-size:12px; }}
.slide-body {{ padding:14px 15px; }}
.change {{ border-left:4px solid #b2ccff; background:#fbfdff; padding:12px 12px; border-radius:10px; margin:10px 0; }}
.kind {{ font-weight:700; margin-bottom:8px; color:#344054; }}
.row {{ display:grid; grid-template-columns:46px 1fr; gap:8px; align-items:start; margin:6px 0; }}
.label {{ font-weight:700; font-size:13px; text-align:center; border-radius:6px; padding:3px 0; }}
.label.old {{ background:#fee4e2; color:#912018; }} .label.new {{ background:#dcfae6; color:#05603a; }}
.text {{ border:1px solid var(--border); border-radius:8px; padding:8px 10px; background:#fff; white-space:pre-wrap; }}
.oldbox {{ background:#fffbfa; }} .newbox {{ background:#f6fef9; }}
mark {{ padding:1px 3px; border-radius:4px; }} mark.del {{ background:#fecdca; color:#7a271a; text-decoration:line-through; }} mark.ins {{ background:#abefc6; color:#054f31; }}
.full {{ color:var(--muted); font-size:13px; margin-top:8px; }}
.muted,.empty {{ color:var(--muted); }}
footer {{ color:var(--muted); font-size:12px; padding:12px 24px 28px; text-align:center; }}
</style>
</head>
<body>
<header>
  <h1>PPT Diff Report</h1>
  <div class="meta">From: {_h(Path(diff['from_file']).name)} &nbsp;→&nbsp; To: {_h(Path(diff['to_file']).name)}<br>Generated at: {_h(diff.get('generated_at'))} · Schema: {_h(diff.get('schema_version'))} · Text mode: {_h(diff.get('settings',{}).get('text_diff_mode'))}</div>
</header>
<main>
  <div class="cards">
    {card('Old slides','old_slide_count')}{card('New slides','new_slide_count')}{card('Added','added','add')}{card('Removed','removed','rem')}{card('Moved','moved','mov')}{card('Modified','modified','mod')}
  </div>
  <section><h2>Added Slides</h2><ul>{added}</ul></section>
  <section><h2>Removed Slides</h2><ul>{removed}</ul></section>
  <section><h2>Moved Slides</h2><ul>{moved}</ul></section>
  <section><h2>Modified Slides</h2>{mods_html}</section>
</main>
<footer>Generated locally by PPT Diff Tool v0.9. JSON and Markdown reports are saved in the same output folder.</footer>
</body>
</html>'''

def write_outputs(diff:Dict,out:Path):
    out.mkdir(parents=True,exist_ok=True)
    base=safe_filename(f"{Path(diff['from_file']).stem}__TO__{Path(diff['to_file']).stem}")
    jp=out/f'{base}.diff.json'; mp=out/f'{base}.diff.md'; hp=out/f'{base}.diff.html'
    jp.write_text(json.dumps(diff,ensure_ascii=False,indent=2),encoding='utf-8')
    mp.write_text(render_markdown(diff),encoding='utf-8')
    hp.write_text(render_html(diff),encoding='utf-8')
    return mp,jp,hp

def run_compare(old_p:Path,new_p:Path,out:Path,detect_format:bool=False):
    diff=build_diff(old_p,new_p,detect_format=detect_format); mp,jp,hp=write_outputs(diff,out); return diff,mp,jp,hp

def launch_ui():
    import tkinter as tk
    from tkinter import filedialog,messagebox
    from tkinter.scrolledtext import ScrolledText
    root=tk.Tk(); root.title('PPT Diff Tool v0.9'); root.geometry('860x610')
    old_var=tk.StringVar(); new_var=tk.StringVar(); out_var=tk.StringVar(value=str(Path.cwd()/'diff_output')); fmt_var=tk.BooleanVar(value=False)
    def browse_old():
        p=filedialog.askopenfilename(title='选择旧版 PPT',filetypes=[('PowerPoint','*.pptx')])
        if p: old_var.set(p); out_var.set(str(Path(p).parent/'diff_output'))
    def browse_new():
        p=filedialog.askopenfilename(title='选择新版 PPT',filetypes=[('PowerPoint','*.pptx')])
        if p: new_var.set(p); out_var.set(str(Path(p).parent/'diff_output'))
    def browse_out():
        p=filedialog.askdirectory(title='选择输出文件夹')
        if p: out_var.set(p)
    def log(msg): output.insert(tk.END,msg+'\n'); output.see(tk.END); root.update_idletasks()
    def open_path(p):
        if not p: return
        try:
            if sys.platform=='darwin': os.system(f'open "{p}"')
            elif os.name=='nt': os.startfile(str(p))
            else: os.system(f'xdg-open "{p}"')
        except Exception as e: messagebox.showerror('打开失败',str(e))
    last_html={'path':None}; last_md={'path':None}; last_out={'path':None}
    def finish_ok(s,mp,jp,hp,outp):
        log('\n完成。'); log(f'HTML：{hp}'); log(f'Markdown：{mp}'); log(f'JSON：{jp}'); log('\nSummary:')
        for k,v in s.items(): log(f'  {k}: {v}')
        last_html['path']=hp; last_md['path']=mp; last_out['path']=outp; btn_html.config(state=tk.NORMAL); btn_report.config(state=tk.NORMAL); btn_folder.config(state=tk.NORMAL); btn_compare.config(state=tk.NORMAL)
    def finish_err(tb):
        log('\n出错：'); log(tb); messagebox.showerror('比较失败',tb[-1500:]); btn_compare.config(state=tk.NORMAL)
    def compare():
        oldp=Path(old_var.get().strip()); newp=Path(new_var.get().strip()); outp=Path(out_var.get().strip())
        if not oldp.exists(): messagebox.showerror('错误','旧版 PPT 不存在。'); return
        if not newp.exists(): messagebox.showerror('错误','新版 PPT 不存在。'); return
        btn_compare.config(state=tk.DISABLED); output.delete('1.0',tk.END)
        log('开始比较...'); log(f'旧版：{oldp.name}'); log(f'新版：{newp.name}'); log(f'输出：{outp}'); log(f'检测格式/布局变化：{fmt_var.get()}'); log('文本差异模式：短句/分句')
        def worker():
            try:
                diff,mp,jp,hp=run_compare(oldp,newp,outp,detect_format=fmt_var.get())
                root.after(0,lambda: finish_ok(diff['summary'],mp,jp,hp,outp))
            except Exception:
                tb=traceback.format_exc(); root.after(0,lambda: finish_err(tb))
        threading.Thread(target=worker,daemon=True).start()
    frm=tk.Frame(root,padx=14,pady=14); frm.pack(fill=tk.BOTH,expand=True)
    tk.Label(frm,text='旧版 PPT（File A）').grid(row=0,column=0,sticky='w')
    tk.Entry(frm,textvariable=old_var,width=86).grid(row=1,column=0,sticky='we',padx=(0,8)); tk.Button(frm,text='选择...',command=browse_old).grid(row=1,column=1)
    tk.Label(frm,text='新版 PPT（File B）').grid(row=2,column=0,sticky='w',pady=(10,0))
    tk.Entry(frm,textvariable=new_var,width=86).grid(row=3,column=0,sticky='we',padx=(0,8)); tk.Button(frm,text='选择...',command=browse_new).grid(row=3,column=1)
    tk.Label(frm,text='输出文件夹').grid(row=4,column=0,sticky='w',pady=(10,0))
    tk.Entry(frm,textvariable=out_var,width=86).grid(row=5,column=0,sticky='we',padx=(0,8)); tk.Button(frm,text='选择...',command=browse_out).grid(row=5,column=1)
    tk.Checkbutton(frm,text='检测格式/布局变化（更敏感；默认关闭。需要检查字号、位置、颜色时再打开）',variable=fmt_var).grid(row=6,column=0,sticky='w',pady=(12,0))
    btns=tk.Frame(frm); btns.grid(row=7,column=0,columnspan=2,sticky='w',pady=(12,8))
    btn_compare=tk.Button(btns,text='比较两个 PPT',width=16,command=compare); btn_compare.pack(side=tk.LEFT)
    btn_html=tk.Button(btns,text='打开 HTML 报告',width=16,state=tk.DISABLED,command=lambda:open_path(last_html['path'])); btn_html.pack(side=tk.LEFT,padx=(8,0))
    btn_report=tk.Button(btns,text='打开 Markdown',width=14,state=tk.DISABLED,command=lambda:open_path(last_md['path'])); btn_report.pack(side=tk.LEFT,padx=(8,0))
    btn_folder=tk.Button(btns,text='打开输出文件夹',width=16,state=tk.DISABLED,command=lambda:open_path(last_out['path'])); btn_folder.pack(side=tk.LEFT,padx=(8,0))
    output=ScrolledText(frm,height=18); output.grid(row=8,column=0,columnspan=2,sticky='nsew')
    frm.columnconfigure(0,weight=1); frm.rowconfigure(8,weight=1)
    log('请选择旧版 PPT、新版 PPT，然后点击“比较两个 PPT”。'); log('v0.9：新增 HTML 可视化报告；文本差异仍按短句/分句展示。')
    root.mainloop()

def main(argv=None)->int:
    p=argparse.ArgumentParser(description='Compare two .pptx files and generate Markdown/JSON diff reports.')
    p.add_argument('old_pptx',nargs='?',type=Path); p.add_argument('new_pptx',nargs='?',type=Path)
    p.add_argument('-o','--output-dir',type=Path,default=Path('diff_output'))
    p.add_argument('--detect-format',action='store_true',help='Detect format/layout XML changes')
    p.add_argument('--ui',action='store_true',help='Launch local GUI')
    args=p.parse_args(argv)
    if args.ui or (args.old_pptx is None and args.new_pptx is None): launch_ui(); return 0
    if args.old_pptx is None or args.new_pptx is None: p.error('old_pptx and new_pptx are required unless --ui is used.')
    try:
        diff,mp,jp,hp=run_compare(args.old_pptx,args.new_pptx,args.output_dir,detect_format=args.detect_format)
        print('Done.'); print(f'HTML:     {hp}'); print(f'Markdown: {mp}'); print(f'JSON:     {jp}\n'); print('Summary:')
        for k,v in diff['summary'].items(): print(f'  {k}: {v}')
        return 0
    except Exception as e:
        print(f'Error: {e}',file=sys.stderr); return 1
if __name__=='__main__': raise SystemExit(main())
