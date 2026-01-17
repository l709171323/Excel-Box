import os
import comtypes.client
import threading

def ppt_to_pdf(input_path, output_path=None, logger=None):
    """
    将单个 PPT/PPTX 文件转换为 PDF
    """
    if not output_path:
        output_path = os.path.splitext(input_path)[0] + ".pdf"
    
    # 提前获取绝对路径，comtypes 需要绝对路径
    input_path = os.path.abspath(input_path)
    output_path = os.path.abspath(output_path)
    
    if logger:
        logger(f"正在转换: {os.path.basename(input_path)} -> {os.path.basename(output_path)}")
    
    powerpoint = None
    try:
        # 初始化 PowerPoint 应用程序
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        # 设置为不可见以加快速度（有些版本可能不支持）
        try:
            powerpoint.Visible = 1 # 1 为 MsoTrue
        except:
            pass
            
        # 打开演示文稿
        # ReadOnly=True, Untitled=False, WithWindow=False
        deck = powerpoint.Presentations.Open(input_path, WithWindow=False)
        
        # 另存为 PDF (32 是 ppSaveAsPDF)
        deck.SaveAs(output_path, 32)
        deck.Close()
        
        if logger:
            logger(f"✅ 转换成功: {os.path.basename(output_path)}")
        return True, output_path
    except Exception as e:
        if logger:
            logger(f"❌ 转换失败: {str(e)}")
        return False, str(e)
    finally:
        if powerpoint:
            try:
                powerpoint.Quit()
            except:
                pass

def batch_ppt_to_pdf(file_paths, output_dir=None, logger=None):
    """
    批量转换 PPT 文件
    """
    success_count = 0
    fail_count = 0
    results = []
    
    # 在批量处理时，为了效率，我们尽量复用一个 PowerPoint 实例
    powerpoint = None
    try:
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        # powerpoint.Visible = 1
        
        for path in file_paths:
            try:
                if not os.path.exists(path):
                    if logger: logger(f"跳过不存在的文件: {path}")
                    fail_count += 1
                    continue
                
                input_path = os.path.abspath(path)
                filename = os.path.basename(input_path)
                out_name = os.path.splitext(filename)[0] + ".pdf"
                
                if output_dir:
                    out_path = os.path.join(os.path.abspath(output_dir), out_name)
                else:
                    out_path = os.path.splitext(input_path)[0] + ".pdf"
                
                if logger: logger(f"正在处理: {filename}...")
                
                # Open: FileName, ReadOnly, Untitled, WithWindow
                deck = powerpoint.Presentations.Open(input_path, WithWindow=False)
                deck.SaveAs(out_path, 32)
                deck.Close()
                
                if logger: logger(f"✅ 完成: {out_name}")
                success_count += 1
                results.append(out_path)
            except Exception as e:
                if logger: logger(f"❌ 转换 {os.path.basename(path)} 失败: {str(e)}")
                fail_count += 1
                
    except Exception as e:
        if logger: logger(f"🔴 PowerPoint 启动失败: {str(e)}")
        raise e
    finally:
        if powerpoint:
            try:
                powerpoint.Quit()
            except:
                pass
                
    return {
        "success": success_count,
        "fail": fail_count,
        "files": results
    }
