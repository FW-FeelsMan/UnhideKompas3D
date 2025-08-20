using System;
using System.IO;
using System.Runtime.InteropServices;
using Kompas6API5;
using Kompas6Constants;
using Kompas6Constants3D;

namespace UnhideKompas3D
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Console.Write("Папка с моделями (*.m3d, *.a3d): ");
            var root = (Console.ReadLine() ?? "").Trim().Trim('"');
            if (string.IsNullOrWhiteSpace(root) || !Directory.Exists(root))
            {
                Console.WriteLine("Папка не найдена.");
                return;
            }

            Console.Write("Рекурсивно по подпапкам? [Y/N]: ");
            bool recursive = ReadYesNo(true);

            KompasObject kompas = null;
            try
            {
                kompas = (KompasObject)Activator.CreateInstance(
                    Type.GetTypeFromProgID("KOMPAS.Application.5", throwOnError: true));
                kompas.Visible = false;

                var files = Directory.GetFiles(
                    root, "*.*",
                    recursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly);

                int total = 0, ok = 0, changedTotal = 0, err = 0;

                foreach (var f in files)
                {
                    var ext = Path.GetExtension(f).ToLowerInvariant();
                    if (ext != ".m3d" && ext != ".a3d") continue;

                    total++;
                    Console.WriteLine($"[{total}] {f}");

                    try
                    {
                        int changed = ProcessOneModel(kompas, f);
                        changedTotal += changed;
                        ok++;
                        Console.WriteLine(changed > 0
                            ? $"   Снято скрытий: {changed}. Сохранено."
                            : "   Нечего менять.");
                    }
                    catch (COMException ex)
                    {
                        err++;
                        Console.WriteLine($"   COM ошибка 0x{ex.HResult:X8}: {ex.Message}");
                    }
                    catch (Exception ex)
                    {
                        err++;
                        Console.WriteLine($"   Ошибка: {ex.Message}");
                    }
                }

                Console.WriteLine();
                Console.WriteLine($"Готово. Обработано: {total}, успешно: {ok}, скрытий снято: {changedTotal}, ошибок: {err}");
            }
            finally
            {
                TryQuit(kompas);
                SafeRelease(ref kompas);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }

            Console.WriteLine("Нажмите Enter...");
            Console.ReadLine();
        }

        // Открыть модель, пройти ВСЕ коллекции сущностей детали и снять скрытие
        static int ProcessOneModel(KompasObject kompas, string path)
        {
            var doc = (ksDocument3D)kompas.Document3D();
            if (doc == null) throw new Exception("Не удалось создать Document3D.");

            // readOnly = false → чтобы сохранить
            if (!doc.Open(path, false)) throw new Exception("Не удалось открыть документ.");

            int changed = 0;
            try
            {
                var part = (ksPart)doc.GetPart((int)Part_Type.pTop_Part);
                if (part == null) throw new Exception("Top Part не получен.");

                // 1) Брут-форс: перебираем ВСЕ типы коллекций (диапазон берём с запасом)
                // В разных версиях SDK набор Obj3dType отличается, это безопаснее всего.
                changed += UnhideEverythingInPart(part);

                // 2) Перестроение и сохранение документа (универсальные вызовы)
                TryCall(doc, "RebuildModel");
                TryCall(doc, "RebuildDocument");
                TryCall(doc, "Save");
            }
            finally
            {
                // Закрыть документ: Close или CloseDocument (что есть)
                TryCall(doc, "Close", (int)DocumentCloseOptions.kdDoNotSaveChanges);
                TryCall(doc, "CloseDocument", (int)DocumentCloseOptions.kdDoNotSaveChanges);
                SafeRelease(ref doc);
            }

            return changed;
        }

        // Перебор всех коллекций и снятие скрытия как у Feature, так и у сущностей без Feature
        static int UnhideEverythingInPart(ksPart part)
        {
            int changed = 0;

            // типы 1..400 — с запасом; лишние просто вернут null/ошибку — игнорируем.
            for (short typeId = 1; typeId <= 400; typeId++)
            {
                ksEntityCollection col = null;
                try { col = (ksEntityCollection)part.EntityCollection(typeId); }
                catch { col = null; }

                if (col == null) continue;

                ksEntity e = null;
                try { e = col.First(); }
                catch { e = null; }

                while (e != null)
                {
                    // сначала пробуем через Feature
                    try
                    {
                        var feat = e.GetFeature();
                        if (feat != null)
                        {
                            if (feat.hidden != 0)
                            {
                                feat.hidden = 0;
                                feat.Update();
                                changed++;
                            }
                        }
                        else
                        {
                            // у некоторых «операций без истории» feature == null, пробуем у сущности скрытие
                            changed += TryUnhideEntityDynamic(e) ? 1 : 0;
                        }
                    }
                    catch
                    {
                        // если GetFeature бросил — пробуем напрямую
                        changed += TryUnhideEntityDynamic(e) ? 1 : 0;
                    }

                    try { e = col.Next(); }
                    catch { break; }
                }
            }

            return changed;
        }

        // Снять скрытие у сущности через dynamic (если есть .hidden/.Visible)
        static bool TryUnhideEntityDynamic(ksEntity e)
        {
            try
            {
                dynamic de = e;
                // разные интеропы по-разному называют свойство
                // 1) hidden: 0/1
                try
                {
                    int h = 0;
                    try { h = (int)de.hidden; } catch { /* нет свойства */ }
                    de.hidden = 0;
                    de.Update();
                    if (h != 0) return true;
                }
                catch { /* игнор */ }

                // 2) Visible: bool
                try
                {
                    bool vis = true;
                    try { vis = (bool)de.Visible; } catch { /* нет свойства */ }
                    de.Visible = true;
                    de.Update();
                    if (!vis) return true;
                }
                catch { /* игнор */ }
            }
            catch { /* игнор */ }
            return false;
        }

        // Универсальные вызовы методов документа (Close/CloseDocument, RebuildModel/RebuildDocument, Save)
        static void TryCall(object obj, string method, params object[] args)
        {
            if (obj == null) return;
            try
            {
                var mi = obj.GetType().GetMethod(method);
                if (mi != null) { mi.Invoke(obj, args); return; }
                dynamic d = obj;
                try { d.GetType(); } catch { return; } // COM мёртв
                try
                {
                    switch (args?.Length ?? 0)
                    {
                        case 0: d.GetType().InvokeMember(method, System.Reflection.BindingFlags.InvokeMethod, null, d, null); break;
                        default: d.GetType().InvokeMember(method, System.Reflection.BindingFlags.InvokeMethod, null, d, args); break;
                    }
                }
                catch { /* игнор */ }
            }
            catch { /* игнор */ }
        }

        static void TryQuit(KompasObject kompas)
        {
            if (kompas == null) return;
            try { kompas.Quit(); } catch { /* игнор */ }
        }

        static void SafeRelease<T>(ref T comObj) where T : class
        {
            var obj = comObj; comObj = null;
            if (obj == null) return;

            try
            {
                if (Marshal.IsComObject(obj))
                {
                    try { while (Marshal.ReleaseComObject(obj) > 0) { } }
                    catch (InvalidComObjectException) { }
                }
            }
            catch (COMException ex)
            {
                // «Запрашиваемый объект отсутствует» — безопасно игнорируем
                if ((uint)ex.HResult != 0x80010114u) throw;
            }
        }

        static bool ReadYesNo(bool defaultYes)
        {
            var s = (Console.ReadLine() ?? "").Trim().ToLowerInvariant();
            if (s == "y" || s == "yes" || s == "д" || s == "да") return true;
            if (s == "n" || s == "no" || s == "н" || s == "нет") return false;
            return defaultYes;
        }
    }
}
