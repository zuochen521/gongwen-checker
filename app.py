import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from pathlib import Path
import zipfile, xml.etree.ElementTree as ET, re, copy

W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
ET.register_namespace('w', W)
def q(n): return '{%s}%s' % (W,n)
def child(parent, name):
    node=parent.find(q(name))
    if node is None: node=ET.SubElement(parent,q(name))
    return node

def process(path, fix=False):
    path=Path(path)
    with zipfile.ZipFile(path) as z:
        files={n:z.read(n) for n in z.namelist()}
    root=ET.fromstring(files['word/document.xml'])
    findings=[]
    for sect in root.iter(q('sectPr')):
        margins=child(sect,'pgMar') if fix else sect.find(q('pgMar'))
        for key, cm in [('top',3.7),('bottom',3.5),('left',2.8),('right',2.6)]:
            expected=round(cm/2.54*1440)
            actual=margins.get(q(key)) if margins is not None else None
            if actual is None or abs(int(actual)-expected)>3:
                findings.append('页面边距 %s：应为 %.1f 厘米（原值：%s twip）' % (key,cm,actual or '未直接设置'))
                if fix: margins.set(q(key),str(expected))
    first_text=True
    for i,p in enumerate(root.iter(q('p')),1):
        text=''.join(t.text or '' for t in p.iter(q('t'))).strip()
        if not text: continue
        if first_text:
            kind, font, size = '标题', '方正小标宋简体', 44
            first_text=False
        elif re.match(r'^[一二三四五六七八九十百]+、',text):
            kind, font, size = '一级标题', '方正黑体简体', 32
        elif re.match(r'^（[一二三四五六七八九十百]+）',text):
            kind, font, size = '二级标题', '方正楷体简体', 32
        elif text == '附件':
            kind, font, size = '附件标识', '方正黑体简体', 32
        else:
            kind, font, size = '正文或其他要素', '方正仿宋简体', 32
        findings.append('第 %d 段：识别为%s，目标字体 %s、%s。请核对识别结果。' % (i,kind,font,'二号' if size==44 else '三号'))
        if fix:
            for run in p.iter(q('r')):
                if not any((t.text or '').strip() for t in run.iter(q('t'))): continue
                props=child(run,'rPr'); fonts=child(props,'rFonts')
                for attr in ('eastAsia','ascii','hAnsi','cs'): fonts.set(q(attr),font)
                child(props,'sz').set(q('val'),str(size))
                child(props,'szCs').set(q('val'),str(size))
            if kind == '标题': child(child(p,'pPr'),'jc').set(q('val'),'center')
        for match in re.finditer(r'(?<!〔)[\[（(](\d{4})[\]）)](?=\d+号)',text):
            findings.append('第 %d 段：发文字号年份括号疑似有误：%s；请人工核对。' % (i,match.group()))
        if text.startswith('附件：') and re.search(r'[。；;]$',text):
            findings.append('第 %d 段：附件名称末尾可能多了标点，请核对。' % i)
        if re.match(r'^\d{4}年\d{1,2}月\d{1,2}日$',text):
            findings.append('第 %d 段：请核对成文日期是否为领导签发日期。' % i)
    if fix:
        files['word/document.xml']=ET.tostring(root,encoding='utf-8',xml_declaration=True)
        output=path.with_name(path.stem+'_已调整字体和页边距.docx')
        if output.exists(): raise FileExistsError('输出文件已存在，请先移动或重命名：'+str(output))
        with zipfile.ZipFile(output,'w',zipfile.ZIP_DEFLATED) as z:
            for name,data in files.items(): z.writestr(name,data)
        return findings,output
    return findings,None

class App:
    def __init__(self, root):
        self.root=root; self.path=None
        root.title('公文校验系统'); root.geometry('1000x720'); root.minsize(780, 580)
        root.configure(bg='#f2f5f9')
        style=ttk.Style(); style.theme_use('clam')
        style.configure('Primary.TButton', font=('Microsoft YaHei UI', 10, 'bold'), padding=(20, 11), background='#2457a7', foreground='white', borderwidth=0)
        style.map('Primary.TButton', background=[('active','#194785')])
        style.configure('Secondary.TButton', font=('Microsoft YaHei UI', 10), padding=(18, 10), background='#e8eef8', foreground='#234b80', borderwidth=0)
        style.map('Secondary.TButton', background=[('active','#d7e4f5')])
        header=tk.Frame(root,bg='#183b70',padx=30,pady=20); header.pack(fill='x')
        tk.Label(header,text='公文校验系统',font=('Microsoft YaHei UI',20,'bold'),fg='white',bg='#183b70').pack(anchor='w')
        tk.Label(header,text='离线处理  ·  公文格式检查与调整',font=('Microsoft YaHei UI',10),fg='#c6d8f3',bg='#183b70').pack(anchor='w',pady=(4,0))
        body=tk.Frame(root,bg='#f2f5f9',padx=28,pady=20); body.pack(fill='both',expand=True)
        card=tk.Frame(body,bg='white',padx=22,pady=18); card.pack(fill='x')
        tk.Label(card,text='选择待处理文件',font=('Microsoft YaHei UI',13,'bold'),fg='#203450',bg='white').pack(anchor='w')
        row=tk.Frame(card,bg='white'); row.pack(fill='x',pady=(12,0))
        ttk.Button(row,text='选择 Word 文件',style='Secondary.TButton',command=self.select).pack(side='left')
        self.label=tk.Label(row,text='尚未选择文件（支持 .docx）',font=('Microsoft YaHei UI',10),fg='#65758b',bg='white',anchor='w'); self.label.pack(side='left',fill='x',expand=True,padx=15)
        actions=tk.Frame(body,bg='#f2f5f9'); actions.pack(fill='x',pady=16)
        ttk.Button(actions,text='开始审核',style='Primary.TButton',command=self.audit).pack(side='left')
        ttk.Button(actions,text='生成调整后的副本',style='Secondary.TButton',command=self.fix).pack(side='left',padx=10)
        self.status=tk.Label(actions,text='等待选择文件',font=('Microsoft YaHei UI',9),fg='#718197',bg='#f2f5f9'); self.status.pack(side='right')
        results=tk.Frame(body,bg='white',padx=22,pady=18); results.pack(fill='both',expand=True)
        tk.Label(results,text='审核结果',font=('Microsoft YaHei UI',13,'bold'),fg='#203450',bg='white').pack(anchor='w')
        tk.Label(results,text='请检查自动识别的段落类型，并在 Word 中复核生成的副本。',font=('Microsoft YaHei UI',9),fg='#7b899c',bg='white').pack(anchor='w',pady=(4,12))
        self.out=scrolledtext.ScrolledText(results,wrap='word',font=('Microsoft YaHei UI',10),fg='#34445b',bg='#f8fafc',relief='flat',padx=14,pady=12,spacing3=5)
        self.out.pack(fill='both',expand=True)
        self.out.insert('end','欢迎使用公文校验系统。\n\n1. 选择需要处理的 Word 文件。\n2. 点击“开始审核”，查看识别结果。\n3. 点击“生成调整后的副本”，在原文件所在目录查看结果。\n\n文件仅在本机处理，原文件不会被覆盖。')
        tk.Label(root,text='本地离线运行  ·  处理完成后请人工复核字体、附件及分页',font=('Microsoft YaHei UI',9),fg='#7a899b',bg='#f2f5f9',pady=8).pack()
    def select(self):
        f=filedialog.askopenfilename(filetypes=[('Word 文档','*.docx')])
        if f:
            self.path=f; self.label.config(text=f,fg='#263d5d'); self.status.config(text='文件已选择')
    def run(self,fix):
        if not self.path: messagebox.showinfo('提示','请先选择 Word 文件'); return
        try:
            findings,output=process(self.path,fix)
            self.out.delete('1.0','end'); self.out.insert('end','审核完成 · 共 %d 条记录\n\n' % len(findings))
            self.out.insert('end','\n'.join(findings) if findings else '未发现当前规则覆盖的问题。')
            self.out.insert('end','\n\n请注意：字体按段落特征自动识别，可能误判；表格、附件标题、分页和印章等须在 Word 中人工复核。')
            if output:
                self.out.insert('end','\n\n副本已保存：'+str(output)); self.status.config(text='副本已生成')
                messagebox.showinfo('处理完成','调整后的副本已保存至：\n'+str(output))
            else: self.status.config(text='审核完成')
        except Exception as e:
            self.status.config(text='处理失败'); messagebox.showerror('处理失败',str(e))
    def audit(self): self.run(False)
    def fix(self): self.run(True)
if __name__=='__main__':
    root=tk.Tk(); App(root); root.mainloop()


