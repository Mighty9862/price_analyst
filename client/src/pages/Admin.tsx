import { useState, useEffect } from "react";
import { motion } from "framer-motion";
import axios from "axios";
import Button from "./components/ui/Button";
import FileIcon from "./components/ui/icons/FileIcon";
import LockIcon from "./components/ui/icons/LockIcon";
import SuccessIcon from "./components/ui/icons/SuccessIcon";
import ErrorIcon from "./components/ui/icons/ErrorIcon";
import DownloadIcon from "./components/ui/icons/DownloadIcon";
import { useAtom } from "jotai";
import { isAdminAtom } from "@/store/store";
import { useNavigate } from "react-router-dom";

interface UploadResponse {
  success: boolean;
  message: string;
  newRecords: number;
  updatedRecords: number;
  unchangedRecords: number;
  skippedRecords: number;
  processedRecords: number;
  failedRecords: number;
  duplicateExamples: string[];
}

function Admin() {
  const API_URL = import.meta.env.VITE_API_URL;
  const ADMIN_PASSWORD = import.meta.env.VITE_ADMIN_PASSWORD;

  const [isAdmin, setIsAdmin] = useAtom(isAdminAtom);
  const navigate = useNavigate();
  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const [password, setPassword] = useState("");
  const [passwordError, setPasswordError] = useState("");
  const [uploadLoading, setUploadLoading] = useState(false);
  const [downloadTemplateLoading, setDownloadTemplateLoading] = useState(false);
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [uploadResponse, setUploadResponse] = useState<UploadResponse | null>(
    null,
  );
  const [uploadMessage, setUploadMessage] = useState<string>("");
  const [uploadSuccess, setUploadSuccess] = useState<boolean | null>(null);

  // Автоматическая аутентификация если isAdmin = true
  useEffect(() => {
    if (isAdmin) {
      setIsAuthenticated(true);
    }
  }, [isAdmin]);

  // Аутентификация администратора
  const handleLogin = (e: React.FormEvent) => {
    e.preventDefault();

    if (!password.trim()) {
      setPasswordError("Введите пароль");
      return;
    }

    if (password === ADMIN_PASSWORD) {
      setIsAuthenticated(true);
      setPasswordError("");
      setIsAdmin(true);
    } else {
      setPasswordError("Неверный пароль");
    }
  };

  // Загрузка шаблона таблицы поставщиков
  const downloadSuppliersTemplate = async () => {
    setDownloadTemplateLoading(true);
    try {
      const response = await axios.get(
        `${API_URL}/template/download-supplier`,
        {
          responseType: "blob",
        },
      );

      const url = window.URL.createObjectURL(new Blob([response.data]));
      const link = document.createElement("a");
      link.href = url;
      link.setAttribute("download", "suppliers_template.xlsx");
      document.body.appendChild(link);
      link.click();
      link.remove();

      setUploadMessage("Шаблон таблицы поставщиков успешно скачан");
      setUploadSuccess(true);
    } catch (error) {
      console.error("Ошибка загрузки шаблона:", error);
      setUploadMessage("Ошибка при скачивании шаблона");
      setUploadSuccess(false);
    } finally {
      setDownloadTemplateLoading(false);
    }
  };

  // Загрузка данных поставщиков
  const uploadSuppliersData = async () => {
    if (!selectedFile) {
      setUploadMessage("Пожалуйста, выберите файл для загрузки");
      setUploadSuccess(false);
      return;
    }

    setUploadLoading(true);
    setUploadResponse(null);
    try {
      const formData = new FormData();
      formData.append("file", selectedFile);

      const response = await axios.post<UploadResponse>(
        `${API_URL}/data/upload-supplier-data`,
        formData,
        {
          headers: {
            "Content-Type": "multipart/form-data",
          },
        },
      );

      console.log(response.data);
      setUploadResponse(response.data);

      if (response.data.success) {
        setUploadMessage("Данные поставщиков успешно загружены");
        setUploadSuccess(true);
      } else {
        setUploadMessage("При загрузке данных возникли ошибки");
        setUploadSuccess(false);
      }

      setSelectedFile(null);

      const fileInput = document.getElementById(
        "suppliers-file-upload",
      ) as HTMLInputElement;
      if (fileInput) fileInput.value = "";
    } catch (error: any) {
      console.error("Ошибка при загрузке файла:", error);
      setUploadMessage(
        error.response?.data?.message ||
          "Ошибка при загрузке данных поставщиков",
      );
      setUploadSuccess(false);
    } finally {
      setUploadLoading(false);
    }
  };

  // Обработчик выбора файла
  const handleFileSelect = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0] || null;
    setSelectedFile(file);
    setUploadMessage("");
    setUploadSuccess(null);
    setUploadResponse(null);
  };

  // Если не аутентифицирован, показываем форму входа
  if (!isAuthenticated) {
    return (
      <div className="from-blue-25 h-screen w-screen bg-gradient-to-br to-white">
        {/* Форма аутентификации */}
        <div className="flex min-h-screen items-center justify-center">
          <div className="container mx-auto px-4 md:px-6">
            <motion.div
              initial={{ opacity: 0, y: 30 }}
              animate={{ opacity: 1, y: 0 }}
              transition={{ duration: 0.7 }}
              className="mx-auto max-w-md"
            >
              <div className="rounded-2xl border border-blue-100 bg-white p-8 shadow-lg shadow-blue-100/30">
                <div className="mb-6 flex justify-center">
                  <div className="flex h-20 w-20 items-center justify-center rounded-2xl bg-gradient-to-br from-blue-500 to-blue-600 text-white shadow-lg">
                    <LockIcon />
                  </div>
                </div>

                <h1 className="mb-2 text-center text-2xl font-bold text-blue-900">
                  Панель администратора
                </h1>
                <p className="mb-6 text-center text-blue-700">
                  Введите пароль для доступа
                </p>

                <form onSubmit={handleLogin}>
                  <div className="mb-4">
                    <label
                      htmlFor="password"
                      className="mb-2 block text-sm font-medium text-blue-700"
                    >
                      Пароль
                    </label>
                    <input
                      id="password"
                      type="password"
                      value={password}
                      onChange={(e) => setPassword(e.target.value)}
                      className={`w-full rounded-lg border px-4 py-3 transition-colors ${
                        passwordError
                          ? "border-red-300 focus:border-red-500 focus:ring-red-500"
                          : "border-blue-200 focus:border-blue-500 focus:ring-blue-500"
                      }`}
                      placeholder="Введите пароль"
                    />
                    {passwordError && (
                      <p className="mt-2 text-sm text-red-600">
                        {passwordError}
                      </p>
                    )}
                  </div>

                  <Button className="w-full">Войти</Button>
                </form>
              </div>
            </motion.div>
          </div>
        </div>
      </div>
    );
  }

  // Если аутентифицирован, показываем панель администратора
  return (
    <div className="from-blue-25 h-screen w-screen bg-gradient-to-br to-white">
      {/* Основной контент */}
      <div className="container mx-auto px-4 py-8 md:px-6">
        <motion.div
          initial={{ opacity: 0, y: 30 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.7, delay: 0.2 }}
          className="mx-auto max-w-4xl"
        >
          {/* Заголовок */}
          <div className="mb-8 text-center">
            <motion.h1
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              transition={{ duration: 0.7, delay: 0.3 }}
              className="mb-4 text-4xl font-bold text-blue-900 md:text-5xl"
            >
              Панель администратора
            </motion.h1>
            <motion.p
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              transition={{ duration: 0.7, delay: 0.4 }}
              className="mb-6 text-lg text-blue-700 md:text-xl"
            >
              Управление данными поставщиков
            </motion.p>
            <Button onClick={() => navigate("/")}>
              Перейти на главную страницу
            </Button>
          </div>

          {/* Карточка загрузки данных */}
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.5, delay: 0.5 }}
            className="rounded-2xl border border-blue-100 bg-white p-8 shadow-lg shadow-blue-100/30"
          >
            <div className="mb-6">
              <h2 className="text-2xl font-semibold text-blue-900">
                Загрузка данных поставщиков
              </h2>
              <p className="mt-2 text-blue-700">
                Загрузите Excel файл с данными поставщиков для обновления базы
                данных
              </p>
            </div>

            {/* Кнопка скачивания шаблона */}
            <div className="mb-6">
              <div className="flex items-center justify-between">
                <div>
                  <h3 className="text-lg font-semibold text-blue-900">
                    Шаблон таблицы
                  </h3>
                  <p className="text-sm text-blue-600">
                    Скачайте готовый шаблон для заполнения
                  </p>
                </div>
                <Button
                  onClick={downloadSuppliersTemplate}
                  disabled={downloadTemplateLoading}
                  className="flex items-center gap-2"
                >
                  {downloadTemplateLoading ? (
                    <span className="flex items-center">
                      <svg
                        className="mr-2 h-4 w-4 animate-spin"
                        fill="none"
                        viewBox="0 0 24 24"
                      >
                        <circle
                          className="opacity-25"
                          cx="12"
                          cy="12"
                          r="10"
                          stroke="currentColor"
                          strokeWidth="4"
                        ></circle>
                        <path
                          className="opacity-75"
                          fill="currentColor"
                          d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"
                        ></path>
                      </svg>
                      Скачивание...
                    </span>
                  ) : (
                    <span className="flex items-center">
                      <DownloadIcon />
                      <span>Скачать шаблон</span>
                    </span>
                  )}
                </Button>
              </div>
            </div>

            <div className="mb-6 rounded-lg border border-blue-200 bg-blue-50 p-4">
              <h3 className="mb-2 font-semibold text-blue-900">
                Требования к файлу:
              </h3>
              <ul className="list-disc space-y-1 pl-5 text-sm text-blue-700">
                <li>Формат: Excel (.xlsx, .xls)</li>
                <li>
                  Должен содержать колонки: штрих-код, наименование товара,
                  поставщик, цена, количество
                </li>
                <li>Первая строка должна содержать заголовки колонок</li>
                <li>Данные должны быть корректно отформатированы</li>
              </ul>
            </div>

            <div className="mb-4">
              <label
                htmlFor="suppliers-file-upload"
                className="mb-3 flex items-center text-sm font-medium text-blue-700"
              >
                <FileIcon />
                <span className="ml-2">
                  Выберите файл с данными поставщиков:
                </span>
              </label>
              <input
                id="suppliers-file-upload"
                type="file"
                accept=".xlsx,.xls"
                onChange={handleFileSelect}
                className="block w-full text-sm text-blue-700 transition-colors file:mr-4 file:rounded-lg file:border-0 file:bg-blue-600 file:px-4 file:py-3 file:text-sm file:font-medium file:text-white hover:file:bg-blue-700"
              />
            </div>

            {selectedFile && (
              <motion.div
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                className="mb-4 rounded-lg border border-blue-200 bg-blue-50 p-3"
              >
                <p className="flex items-center text-sm text-blue-700">
                  <FileIcon />
                  <span className="ml-2">
                    Выбран файл: <strong>{selectedFile.name}</strong>
                  </span>
                </p>
              </motion.div>
            )}

            <Button
              onClick={uploadSuppliersData}
              disabled={uploadLoading || !selectedFile}
              className="w-full"
            >
              {uploadLoading ? (
                <span className="flex items-center justify-center">
                  <svg
                    className="mr-2 h-4 w-4 animate-spin"
                    fill="none"
                    viewBox="0 0 24 24"
                  >
                    <circle
                      className="opacity-25"
                      cx="12"
                      cy="12"
                      r="10"
                      stroke="currentColor"
                      strokeWidth="4"
                    ></circle>
                    <path
                      className="opacity-75"
                      fill="currentColor"
                      d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"
                    ></path>
                  </svg>
                  Загрузка...
                </span>
              ) : (
                <span className="flex items-center justify-center">
                  <FileIcon />
                  <span className="ml-2">Загрузить данные поставщиков</span>
                </span>
              )}
            </Button>
          </motion.div>

          {/* Сообщение о статусе */}
          {uploadMessage && (
            <motion.div
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              className={`mt-6 rounded-xl p-4 ${
                uploadSuccess === true
                  ? "border border-green-200 bg-green-50 text-green-800"
                  : uploadSuccess === false
                    ? "border border-red-200 bg-red-50 text-red-800"
                    : "border border-blue-200 bg-blue-50 text-blue-800"
              }`}
            >
              <div className="flex items-center">
                {uploadSuccess === true ? <SuccessIcon /> : <ErrorIcon />}
                <span className="ml-2 font-medium">{uploadMessage}</span>
              </div>
            </motion.div>
          )}

          {/* Детальный ответ от сервера */}
          {uploadResponse && (
            <motion.div
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              transition={{ duration: 0.5, delay: 0.3 }}
              className="mt-6 rounded-2xl border border-blue-100 bg-white p-6 shadow-lg shadow-blue-100/30"
            >
              <h3 className="mb-4 text-xl font-semibold text-blue-900">
                Результат обработки
              </h3>

              {/* Основная статистика */}
              <div className="mb-6 grid grid-cols-2 gap-4 md:grid-cols-4">
                <div className="rounded-lg bg-green-100 p-4 text-center text-green-800">
                  <div className="text-2xl font-bold">
                    {uploadResponse.newRecords}
                  </div>
                  <div className="text-sm">Новых записей</div>
                </div>
                <div className="rounded-lg bg-blue-100 p-4 text-center text-blue-800">
                  <div className="text-2xl font-bold">
                    {uploadResponse.updatedRecords}
                  </div>
                  <div className="text-sm">Обновлено</div>
                </div>
                <div className="rounded-lg bg-gray-100 p-4 text-center text-gray-800">
                  <div className="text-2xl font-bold">
                    {uploadResponse.unchangedRecords}
                  </div>
                  <div className="text-sm">Без изменений</div>
                </div>
                <div className="rounded-lg bg-amber-100 p-4 text-center text-amber-800">
                  <div className="text-2xl font-bold">
                    {uploadResponse.skippedRecords}
                  </div>
                  <div className="text-sm">Пропущено дубликатов</div>
                </div>
              </div>

              {/* Дополнительная статистика */}
              <div className="mb-6 grid grid-cols-2 gap-4">
                <div
                  className={`rounded-lg p-4 text-center ${
                    uploadResponse.processedRecords > 0
                      ? "bg-green-100 text-green-800"
                      : "bg-blue-100 text-blue-800"
                  }`}
                >
                  <div className="text-2xl font-bold">
                    {uploadResponse.processedRecords}
                  </div>
                  <div className="text-sm">Обработано записей</div>
                </div>
                <div
                  className={`rounded-lg p-4 text-center ${
                    uploadResponse.failedRecords > 0
                      ? "bg-red-100 text-red-800"
                      : "bg-green-100 text-green-800"
                  }`}
                >
                  <div className="text-2xl font-bold">
                    {uploadResponse.failedRecords}
                  </div>
                  <div className="text-sm">Ошибок</div>
                </div>
              </div>
              {/* Примеры дубликатов */}
              {uploadResponse.duplicateExamples.length > 0 && (
                <div>
                  <h4 className="mb-3 font-semibold text-blue-900">
                    Примеры дубликатов (
                    {uploadResponse.duplicateExamples.length}):
                  </h4>
                  <div className="max-h-60 overflow-y-auto rounded-lg border border-blue-200">
                    <div className="divide-y divide-blue-100">
                      {uploadResponse.duplicateExamples.map(
                        (example, index) => (
                          <div key={index} className="p-3">
                            <p className="text-sm text-blue-700">{example}</p>
                          </div>
                        ),
                      )}
                    </div>
                  </div>
                  <p className="mt-2 text-sm text-blue-600">
                    Показаны первые {uploadResponse.duplicateExamples.length}{" "}
                    дубликатов
                  </p>
                </div>
              )}

              {/* Статус операции */}
              <div className="mt-4 flex items-center justify-between border-t border-blue-200 pt-4">
                <span
                  className={`inline-flex items-center rounded-full px-3 py-1 text-sm font-medium ${
                    uploadResponse.success
                      ? "bg-green-100 text-green-800"
                      : "bg-red-100 text-red-800"
                  }`}
                >
                  {uploadResponse.success
                    ? "Данные успешно загружены"
                    : "Обработка завершена с ошибками"}
                </span>
              </div>
            </motion.div>
          )}
        </motion.div>
      </div>
    </div>
  );
}

export default Admin;
