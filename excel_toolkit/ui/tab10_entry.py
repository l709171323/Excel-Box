"""
Tab10 - 录入发货信息功能
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os

from excel_toolkit.warehouse_router import write_inventory, read_inventory


class Tab10EntryMixin:
    """Tab10 录入发货信息 Mixin"""
    
    def create_tab10_entry(self, tab):
        """创建Tab10界面"""
        self.wh10 = {}  # 仓库->州 映射
        self.sku10 = {}  # 仓库->SKU集合 映射

        # 左侧：仓库列表
        left_frame = ttk.LabelFrame(tab, text="仓库列表", style="Section.TLabelframe")
        left_frame.pack(side='left', fill='both', expand=True, padx=5, pady=5)
        
        # 仓库表格
        columns = ('仓库名称', '州')
        self.tree10 = ttk.Treeview(left_frame, columns=columns, show='headings', height=10)
        self.tree10.heading('仓库名称', text='仓库名称')
        self.tree10.heading('州', text='州')
        self.tree10.column('仓库名称', width=150)
        self.tree10.column('州', width=60)
        self.tree10.pack(fill='both', expand=True, padx=5, pady=5)
        self.tree10.bind('<<TreeviewSelect>>', self._on_wh_select10)
        
        # 添加仓库按钮
        btn_frame1 = ttk.Frame(left_frame)
        btn_frame1.pack(fill='x', padx=5, pady=5)
        ttk.Button(btn_frame1, text="➕ 添加仓库", 
                  command=self._add_warehouse10).pack(side='left', padx=2)
        ttk.Button(btn_frame1, text="➖ 删除仓库", 
                  command=self._del_warehouse10).pack(side='left', padx=2)

        # 右侧：SKU列表
        right_frame = ttk.LabelFrame(tab, text="SKU列表（选中仓库）", style="Section.TLabelframe")
        right_frame.pack(side='left', fill='both', expand=True, padx=5, pady=5)
        
        self.list10 = tk.Listbox(right_frame, height=12, selectmode='extended')
        self.list10.pack(fill='both', expand=True, padx=5, pady=5)
        
        # 添加SKU按钮
        btn_frame2 = ttk.Frame(right_frame)
        btn_frame2.pack(fill='x', padx=5, pady=5)
        ttk.Button(btn_frame2, text="➕ 添加SKU", 
                  command=self._add_sku10).pack(side='left', padx=2)
        ttk.Button(btn_frame2, text="➖ 删除SKU", 
                  command=self._del_sku10).pack(side='left', padx=2)

        # 底部操作区
        bottom_frame = ttk.Frame(tab)
        bottom_frame.pack(side='bottom', fill='x', padx=5, pady=10)
        
        # 文件路径显示
        file_frame = ttk.Frame(bottom_frame)
        file_frame.pack(fill='x', pady=(0, 5))
        ttk.Label(file_frame, text="当前库存文件:").pack(side='left', padx=(0, 5))
        ttk.Label(file_frame, textvariable=self.inv10_var, 
                 foreground='blue').pack(side='left', padx=(0, 5))
        
        ttk.Button(bottom_frame, text="📂 导入库存文件", 
                  command=self._import_inventory10).pack(side='left', padx=5)
        ttk.Button(bottom_frame, text="💾 保存库存文件", 
                  command=self._save_inventory10, 
                  style='Accent.TButton').pack(side='left', padx=5)
        ttk.Button(bottom_frame, text="🗑️ 清空所有", 
                  command=self._clear_all10).pack(side='left', padx=5)
        
        # 数据库同步按钮
        ttk.Label(bottom_frame, text="|", font=('Segoe UI', 12)).pack(side='left', padx=8)
        ttk.Button(bottom_frame, text="📤 保存到数据库", 
                  command=self._save_to_database10,
                  style='Secondary.TButton').pack(side='left', padx=5)
        ttk.Button(bottom_frame, text="📥 从数据库加载", 
                  command=self._load_from_database10,
                  style='Secondary.TButton').pack(side='left', padx=5)
        
        self.logger10, clear_log10 = self.create_log_widget(tab)
        
        # Tab10创建完成后，检查是否需要自动加载数据
        self.master.after(100, self._auto_load_tab10_data)
    
    def _auto_load_tab10_data(self):
        """Tab10创建完成后自动加载数据"""
        try:
            # 优先从数据库加载
            if hasattr(self, 'wh10') and self.wh10:
                # 如果已经有数据（可能是从数据库预加载的），直接更新UI
                for w, st in sorted(self.wh10.items()):
                    self.tree10.insert('', 'end', values=(w, st or ''))
                self.logger10(f"✅ 自动加载了 {len(self.wh10)} 个仓库（来自数据库）")
                return
            
            # 从文件加载
            if hasattr(self, 'inv10_var'):
                inv_path = self.inv10_var.get()
                if inv_path and inv_path != "未选择库存文件" and inv_path != "[数据库]" and os.path.exists(inv_path):
                    sku_by_wh, wh_state = read_inventory(inv_path, logger=lambda x: None)
                    
                    # 更新数据
                    self.wh10 = {str(k): str(v) if v else '' for k, v in wh_state.items()}
                    self.sku10 = {str(k): set(v) for k, v in sku_by_wh.items()}
                    
                    # 更新UI
                    for w, st in sorted(self.wh10.items()):
                        self.tree10.insert('', 'end', values=(w, st or ''))
                    
                    self.logger10(f"✅ 自动加载了 {len(self.wh10)} 个仓库（来自文件: {os.path.basename(inv_path)}）")
                    
                    # 同步到Tab9
                    if hasattr(self, 'inv9_var'):
                        self.inv9_var.set(inv_path)
                        
        except Exception as e:
            self.logger10(f"⚠️ 自动加载失败: {e}")

    def _on_wh_select10(self, event=None):
        """仓库选择变化时更新SKU列表"""
        sel = self.tree10.selection()
        if not sel:
            return
        
        item = self.tree10.item(sel[0])
        wh_name = item['values'][0]
        
        self.list10.delete(0, 'end')
        skus = self.sku10.get(wh_name, set())
        for sku in sorted(skus):
            self.list10.insert('end', sku)

    def _add_warehouse10(self):
        """添加仓库"""
        dialog = tk.Toplevel(self.master)
        dialog.title("添加仓库")
        dialog.geometry("300x150")
        dialog.transient(self.master)
        dialog.grab_set()
        
        ttk.Label(dialog, text="仓库名称:").pack(pady=(10, 0))
        name_entry = ttk.Entry(dialog, width=30)
        name_entry.pack(pady=5)
        
        ttk.Label(dialog, text="州（两位缩写）:").pack()
        state_entry = ttk.Entry(dialog, width=10)
        state_entry.pack(pady=5)
        
        def on_ok():
            name = name_entry.get().strip()
            state = state_entry.get().strip().upper()
            if not name:
                messagebox.showwarning("警告", "请输入仓库名称")
                return
            
            self.wh10[name] = state
            self.sku10[name] = set()
            self.tree10.insert('', 'end', values=(name, state))
            self.logger10(f"已添加仓库: {name} ({state})")
            dialog.destroy()
        
        ttk.Button(dialog, text="确定", command=on_ok, 
                  style='Accent.TButton').pack(pady=10)

    def _del_warehouse10(self):
        """删除仓库"""
        sel = self.tree10.selection()
        if not sel:
            messagebox.showwarning("警告", "请先选择要删除的仓库")
            return
        
        item = self.tree10.item(sel[0])
        wh_name = item['values'][0]
        
        if messagebox.askyesno("确认", f"确定要删除仓库 '{wh_name}' 吗？"):
            self.tree10.delete(sel[0])
            self.wh10.pop(wh_name, None)
            self.sku10.pop(wh_name, None)
            self.list10.delete(0, 'end')
            self.logger10(f"已删除仓库: {wh_name}")

    def _add_sku10(self):
        """添加SKU"""
        sel = self.tree10.selection()
        if not sel:
            messagebox.showwarning("警告", "请先选择仓库")
            return
        
        item = self.tree10.item(sel[0])
        wh_name = item['values'][0]
        
        dialog = tk.Toplevel(self.master)
        dialog.title(f"添加SKU到 {wh_name}")
        dialog.geometry("300x150")
        dialog.transient(self.master)
        dialog.grab_set()
        
        ttk.Label(dialog, text="SKU（多个用逗号分隔）:").pack(pady=(10, 0))
        sku_entry = ttk.Entry(dialog, width=30)
        sku_entry.pack(pady=10)
        
        def on_ok():
            skus = [s.strip() for s in sku_entry.get().split(',') if s.strip()]
            if not skus:
                messagebox.showwarning("警告", "请输入SKU")
                return
            
            if wh_name not in self.sku10:
                self.sku10[wh_name] = set()
            
            for sku in skus:
                self.sku10[wh_name].add(sku)
                self.list10.insert('end', sku)
            
            self.logger10(f"已添加 {len(skus)} 个SKU到 {wh_name}")
            dialog.destroy()
        
        ttk.Button(dialog, text="确定", command=on_ok, 
                  style='Accent.TButton').pack(pady=10)

    def _del_sku10(self):
        """删除SKU"""
        sel_wh = self.tree10.selection()
        if not sel_wh:
            messagebox.showwarning("警告", "请先选择仓库")
            return
        
        sel_sku = self.list10.curselection()
        if not sel_sku:
            messagebox.showwarning("警告", "请选择要删除的SKU")
            return
        
        item = self.tree10.item(sel_wh[0])
        wh_name = item['values'][0]
        
        # 倒序删除
        for i in reversed(sel_sku):
            sku = self.list10.get(i)
            self.list10.delete(i)
            self.sku10[wh_name].discard(sku)
        
        self.logger10(f"已删除 {len(sel_sku)} 个SKU")

    def _import_inventory10(self):
        """导入库存文件"""
        path = filedialog.askopenfilename(
            title="选择库存文件",
            filetypes=[("Excel文件", "*.xlsx;*.xlsm;*.xls"), ("所有文件", "*.*")]
        )
        if not path:
            return
        
        try:
            sku_by_wh, wh_state = read_inventory(path, logger=self.logger10)
            
            # 清空现有数据
            for item in self.tree10.get_children():
                self.tree10.delete(item)
            self.list10.delete(0, 'end')
            
            self.wh10 = {str(k): str(v) if v else '' for k, v in wh_state.items()}
            self.sku10 = {str(k): set(v) for k, v in sku_by_wh.items()}
            
            # 更新UI
            for w, st in sorted(self.wh10.items()):
                self.tree10.insert('', 'end', values=(w, st or ''))
            
            # 保存文件路径到持久化变量
            self.inv10_var.set(path)
            self.logger10(f"已导入库存文件: {path}")
            
            # 同步到Tab9
            self.inv9_var.set(path)
            if hasattr(self, '_refresh_block9_from_inventory'):
                self._refresh_block9_from_inventory()
                
        except Exception as e:
            messagebox.showerror("错误", f"导入失败: {e}")
            self.logger10(f"导入失败: {e}")

    def _save_inventory10(self):
        """保存库存文件"""
        if not self.wh10:
            messagebox.showwarning("警告", "没有数据可保存")
            return
        
        path = filedialog.asksaveasfilename(
            title="保存库存文件",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if not path:
            return
        
        try:
            result = write_inventory(path, self.wh10, self.sku10, self.logger10)
            messagebox.showinfo("完成", result)
            
            # 保存文件路径到持久化变量
            self.inv10_var.set(path)
            
            # 同步到Tab9
            self.inv9_var.set(path)
            if hasattr(self, '_refresh_block9_from_inventory'):
                self._refresh_block9_from_inventory()
                
        except Exception as e:
            messagebox.showerror("错误", f"保存失败: {e}")
            self.logger10(f"保存失败: {e}")

    def _clear_all10(self):
        """清空所有数据"""
        if messagebox.askyesno("确认", "确定要清空所有数据吗？"):
            for item in self.tree10.get_children():
                self.tree10.delete(item)
            self.list10.delete(0, 'end')
            self.wh10.clear()
            self.sku10.clear()
            self.logger10("已清空所有数据")
    
    def _save_to_database10(self):
        """保存当前库存到数据库"""
        if not self.wh10:
            messagebox.showwarning("警告", "没有数据可保存")
            return
        
        try:
            from excel_toolkit.db_operations import save_warehouse_inventory
            from excel_toolkit.db_config import get_db_manager
            
            # 检查数据库是否启用
            db = get_db_manager()
            if not db.config.is_enabled():
                messagebox.showwarning("数据库未启用", 
                                     "数据库功能未启用。\n\n"
                                     "请在程序目录下创建 db_config.json 文件：\n"
                                     '{\n  "enabled": true,\n  "type": "sqlite"\n}')
                return
            
            success, msg = save_warehouse_inventory(self.wh10, self.sku10)
            
            if success:
                messagebox.showinfo("成功", msg)
                self.logger10(f"✅ {msg}")
            else:
                messagebox.showerror("失败", msg)
                self.logger10(f"❌ {msg}")
        except Exception as e:
            messagebox.showerror("错误", f"保存到数据库失败: {e}")
            self.logger10(f"❌ 保存到数据库失败: {e}")
    
    def _load_from_database10(self):
        """从数据库加载库存"""
        try:
            from excel_toolkit.db_operations import load_warehouse_inventory
            from excel_toolkit.db_config import get_db_manager
            
            # 检查数据库是否启用
            db = get_db_manager()
            if not db.config.is_enabled():
                messagebox.showwarning("数据库未启用", 
                                     "数据库功能未启用。\n\n"
                                     "请在程序目录下创建 db_config.json 文件：\n"
                                     '{\n  "enabled": true,\n  "type": "sqlite"\n}')
                return
            
            data = load_warehouse_inventory()
            
            if not data:
                messagebox.showwarning("警告", "数据库中没有库存数据")
                return
            
            warehouse_data, sku_data = data
            
            # 清空现有数据
            for item in self.tree10.get_children():
                self.tree10.delete(item)
            self.list10.delete(0, 'end')
            
            self.wh10 = warehouse_data
            self.sku10 = sku_data
            
            # 更新UI
            for w, st in sorted(self.wh10.items()):
                self.tree10.insert('', 'end', values=(w, st or ''))
            
            total_wh = len(self.wh10)
            total_sku = sum(len(skus) for skus in self.sku10.values())
            messagebox.showinfo("成功", f"已加载 {total_wh} 个仓库，{total_sku} 个SKU")
            self.logger10(f"✅ 已从数据库加载 {total_wh} 个仓库")
            
            # 更新文件路径显示为数据库标识
            self.inv10_var.set("[数据库]")
            
            # 同步到Tab9
            if hasattr(self, 'inv9_var'):
                self.inv9_var.set("[数据库]")
                if hasattr(self, '_refresh_block9_from_inventory'):
                    # 使用现有的刷新机制
                    try:
                        from excel_toolkit.warehouse_router import read_inventory
                        # 模拟从数据库读取
                        self.master.after(100, self._refresh_block9_from_inventory)
                    except:
                        pass
                    
        except Exception as e:
            messagebox.showerror("错误", f"从数据库加载失败: {e}")
            self.logger10(f"❌ 从数据库加载失败: {e}")





























