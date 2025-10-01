import { useState } from "react";
import { motion } from "framer-motion";
import axios from "axios";

// Типы данных для ответа от API
interface PriceAnalysisResult {
  barcode: string;
  quantity: number;
  bestSupplierName: string;
  bestSupplierSap: string;
  bestPrice: number;
  productName: string;
  requiresManualProcessing: boolean;
}

// SVG иконки
const DownloadIcon = () => (
  <svg
    className="h-6 w-6"
    viewBox="0 0 24 24"
    fill="none"
    xmlns="http://www.w3.org/2000/svg"
  >
    <path
      d="M21 15V16.2C21 17.8802 21 18.7202 20.673 19.362C20.3854 19.9265 19.9265 20.3854 19.362 20.673C18.7202 21 17.8802 21 16.2 21H7.8C6.11984 21 5.27976 21 4.63803 20.673C4.07354 20.3854 3.6146 19.9265 3.32698 19.362C3 18.7202 3 17.8802 3 16.2V15M17 10L12 15M12 15L7 10M12 15V3"
      stroke="currentColor"
      strokeWidth="2"
      strokeLinecap="round"
      strokeLinejoin="round"
    />
  </svg>
);

const AnalysisIcon = () => (
  <svg
    className="h-6 w-6"
    viewBox="0 0 24 24"
    fill="none"
    xmlns="http://www.w3.org/2000/svg"
  >
    <path
      d="M7.5 14.25V16.5M12 12.75V16.5M16.5 11.25V16.5M4.5 18H19.5C20.7426 18 21.75 16.9926 21.75 15.75V6.75C21.75 5.50736 20.7426 4.5 19.5 4.5H4.5C3.25736 4.5 2.25 5.50736 2.25 6.75V15.75C2.25 16.9926 3.25736 18 4.5 18Z"
      stroke="currentColor"
      strokeWidth="2"
      strokeLinecap="round"
      strokeLinejoin="round"
    />
  </svg>
);

const SuccessIcon = () => (
  <svg
    className="h-5 w-5"
    viewBox="0 0 24 24"
    fill="none"
    xmlns="http://www.w3.org/2000/svg"
  >
    <path
      d="M20 6L9 17L4 12"
      stroke="currentColor"
      strokeWidth="2"
      strokeLinecap="round"
      strokeLinejoin="round"
    />
  </svg>
);

const FileIcon = () => (
  <svg
    className="h-5 w-5"
    viewBox="0 0 24 24"
    fill="none"
    xmlns="http://www.w3.org/2000/svg"
  >
    <path
      d="M14 2.26953V6.40007C14 6.96012 14 7.24015 14.109 7.45406C14.2049 7.64222 14.3578 7.7952 14.546 7.89108C14.7599 8.00007 15.0399 8.00007 15.6 8.00007H19.7305M14 2.26953C13.8219 2.26163 13.6415 2.25372 13.4743 2.26605C13.0753 2.29779 12.6759 2.41152 12.343 2.6527C12.1275 2.81087 11.9397 2.99862 11.5642 3.37414L5.43416 9.50418C4.87411 10.0642 4.59408 10.3442 4.41769 10.681C4.26711 10.971 4.17299 11.2859 4.14026 11.6099C4.10294 11.9834 4.10294 12.3722 4.10294 13.1496V16.8404C4.10294 18.0831 4.10294 18.7044 4.29291 19.1875C4.45691 19.6051 4.7275 19.9705 5.07614 20.246C5.47703 20.5628 6.02248 20.7084 6.75065 20.7712C7.59353 20.8455 8.181 20.8827 8.70412 20.8154C9.09972 20.7654 9.48884 20.6703 9.86284 20.532C10.3989 20.3302 10.8756 20.0086 11.8288 19.3654L13.8597 18.0151C14.2671 17.7445 14.4708 17.6092 14.6445 17.4443C14.7978 17.2987 14.9335 17.1359 15.0485 16.9592C15.1795 16.7583 15.2754 16.5344 15.4673 16.0867L16.219 14.3468C16.5299 13.6085 16.6853 13.2394 16.7254 12.8613C16.7499 12.6328 16.7499 12.4021 16.7254 12.1736C16.6853 11.7955 16.5299 11.4264 16.219 10.6881L15.4673 8.94815C15.2754 8.50048 15.1795 8.27665 15.0485 8.07578C14.9335 7.89904 14.7978 7.7362 14.6445 7.59066C14.4708 7.42573 14.2671 7.29041 13.8597 7.01977L12.6256 6.16154C12.2491 5.91335 12.0609 5.78926 11.8626 5.69889C11.624 5.59095 11.3729 5.51203 11.1148 5.46392C10.892 5.42271 10.6628 5.41718 10.2045 5.40612C9.46171 5.38893 8.86814 5.30238 8.31593 5.0779C7.86371 4.89273 7.45095 4.6226 7.1023 4.28377C6.69048 3.88394 6.38466 3.37129 5.77303 2.34599L14 2.26953Z"
      stroke="currentColor"
      strokeWidth="2"
      strokeLinecap="round"
    />
  </svg>
);

// Анимированная кнопка с голубыми цветами
const AnimatedButton = ({
  children,
  primary = true,
  className = "",
  onClick,
  disabled = false,
}: {
  children: React.ReactNode;
  primary?: boolean;
  className?: string;
  onClick?: () => void;
  disabled?: boolean;
}) => (
  <motion.button
    className={`rounded-lg px-6 py-3 text-lg font-medium transition-colors ${
      primary
        ? "bg-blue-600 text-white shadow-lg shadow-blue-500/30 hover:bg-blue-700"
        : "border-2 border-blue-200 text-blue-700 hover:bg-blue-50"
    } ${disabled ? "cursor-not-allowed opacity-50" : ""} ${className}`}
    whileHover={!disabled ? { scale: 1.05 } : {}}
    whileTap={!disabled ? { scale: 0.95 } : {}}
    initial={{ opacity: 0, y: 10 }}
    animate={{ opacity: 1, y: 0 }}
    onClick={disabled ? undefined : onClick}
    disabled={disabled}
  >
    {children}
  </motion.button>
);

function Home() {
  const API_URL = import.meta.env.VITE_API_URL;
  const [downloadLoading, setDownloadLoading] = useState(false);
  const [uploadLoading, setUploadLoading] = useState(false);
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [uploadMessage, setUploadMessage] = useState<string>("");
  const [uploadSuccess, setUploadSuccess] = useState<boolean | null>(null);
  const [analysisResults, setAnalysisResults] = useState<PriceAnalysisResult[]>(
    [],
  );
  const [showResults, setShowResults] = useState(false);

  // Функция для скачивания шаблона таблицы
  const downloadTemplate = async () => {
    setDownloadLoading(true);
    try {
      const response = await axios.get(`${API_URL}/template/download`, {
        responseType: "blob",
      });

      const url = window.URL.createObjectURL(new Blob([response.data]));
      const link = document.createElement("a");
      link.href = url;
      link.setAttribute("download", "template.xlsx");
      document.body.appendChild(link);
      link.click();
      link.remove();

      setUploadMessage("Шаблон успешно скачан");
      setUploadSuccess(true);
    } catch (error) {
      console.error("Ошибка загрузки:", error);
      setUploadMessage("Ошибка при скачивании шаблона");
      setUploadSuccess(false);
    } finally {
      setDownloadLoading(false);
    }
  };

  // Функция для загрузки заполненной таблицы
  const uploadTable = async () => {
    if (!selectedFile) {
      setUploadMessage("Пожалуйста, выберите файл для загрузки");
      setUploadSuccess(false);
      return;
    }

    setUploadLoading(true);
    try {
      const formData = new FormData();
      formData.append("file", selectedFile);

      const response = await axios.post<PriceAnalysisResult[]>(
        `${API_URL}/data/analyze-prices`,
        formData,
        {
          headers: {
            "Content-Type": "multipart/form-data",
          },
        },
      );

      setUploadMessage("Анализ цен завершен успешно");
      setUploadSuccess(true);

      console.log(response.data);
      if (response.data) {
        setAnalysisResults(response.data);
        setShowResults(true);
      }

      setSelectedFile(null);

      // Сбрасываем input file
      const fileInput = document.getElementById(
        "file-upload",
      ) as HTMLInputElement;
      if (fileInput) fileInput.value = "";
    } catch (error: any) {
      console.error("Ошибка при загрузке файла:", error);
      setUploadMessage(
        error.response?.data?.message || "Ошибка при анализе файла",
      );
      setUploadSuccess(false);
      setShowResults(false);
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
    setShowResults(false);
  };

  // Форматирование цены
  const formatPrice = (price: number) => {
    return new Intl.NumberFormat("ru-RU", {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    }).format(price);
  };

  return (
    <div className="from-blue-25 relative h-screen w-screen bg-gradient-to-br to-white">
      {/* Основной контент */}
      <div className="relative z-10 flex min-h-screen items-center justify-center pt-24 pb-12">
        <div className="container mx-auto px-4 md:px-6">
          <motion.div
            initial={{ opacity: 0, y: 30 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.7, delay: 0.2 }}
            className="mx-auto max-w-6xl"
          >
            {/* Заголовок */}
            <div className="mb-16 text-center">
              <motion.h1
                initial={{ opacity: 0, y: 20 }}
                animate={{ opacity: 1, y: 0 }}
                transition={{ duration: 0.7, delay: 0.3 }}
                className="mb-4 text-4xl font-bold text-blue-900 md:text-5xl"
              >
                Анализ цен поставщиков
              </motion.h1>
              <motion.p
                initial={{ opacity: 0, y: 20 }}
                animate={{ opacity: 1, y: 0 }}
                transition={{ duration: 0.7, delay: 0.4 }}
                className="mx-auto max-w-2xl text-lg text-blue-700 md:text-xl"
              >
                Получите лучшие цены на товары от проверенных поставщиков в
                несколько кликов
              </motion.p>
            </div>

            {/* Карточки действий */}
            <div className="mb-16 grid grid-cols-1 gap-8 lg:grid-cols-2">
              {/* Карточка скачивания шаблона */}
              <motion.div
                initial={{ opacity: 0, x: -30 }}
                animate={{ opacity: 1, x: 0 }}
                transition={{ duration: 0.5, delay: 0.5 }}
                className="group relative rounded-2xl border border-blue-100 bg-white p-8 shadow-lg shadow-blue-100/30 transition-all duration-300 hover:shadow-xl hover:shadow-blue-100/50"
              >
                <div className="mb-6 flex h-20 w-20 items-center justify-center rounded-2xl bg-gradient-to-br from-blue-500 to-blue-600 text-white shadow-lg transition-transform duration-300 group-hover:scale-110">
                  <DownloadIcon />
                </div>
                <h3 className="mb-4 text-2xl font-semibold text-blue-900">
                  Скачать шаблон
                </h3>
                <p className="mb-6 leading-relaxed text-blue-700">
                  Получите готовый шаблон Excel таблицы для заполнения данными о
                  товарах. Простая структура для быстрого заполнения.
                </p>
                <AnimatedButton
                  onClick={downloadTemplate}
                  disabled={downloadLoading}
                  className="w-full"
                >
                  {downloadLoading ? (
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
                      Скачивание...
                    </span>
                  ) : (
                    <span className="flex items-center justify-center">
                      <DownloadIcon />
                      <span className="ml-2">Скачать шаблон</span>
                    </span>
                  )}
                </AnimatedButton>
              </motion.div>

              {/* Карточка загрузки таблицы */}
              <motion.div
                initial={{ opacity: 0, x: 30 }}
                animate={{ opacity: 1, x: 0 }}
                transition={{ duration: 0.5, delay: 0.6 }}
                className="group relative rounded-2xl border border-blue-100 bg-white p-8 shadow-lg shadow-blue-100/30 transition-all duration-300 hover:shadow-xl hover:shadow-blue-100/50"
              >
                <div className="mb-6 flex h-20 w-20 items-center justify-center rounded-2xl bg-gradient-to-br from-cyan-500 to-blue-500 text-white shadow-lg transition-transform duration-300 group-hover:scale-110">
                  <AnalysisIcon />
                </div>
                <h3 className="mb-4 text-2xl font-semibold text-blue-900">
                  Анализ цен
                </h3>
                <p className="mb-6 leading-relaxed text-blue-700">
                  Загрузите заполненную таблицу для автоматического анализа.
                  Система найдет поставщиков с лучшими ценами для каждого
                  товара.
                </p>

                <div className="mb-4">
                  <label
                    htmlFor="file-upload"
                    className="mb-3 flex items-center text-sm font-medium text-blue-700"
                  >
                    <FileIcon />
                    <span className="ml-2">Выберите файл для анализа:</span>
                  </label>
                  <input
                    id="file-upload"
                    type="file"
                    accept=".xlsx,.xls,.csv"
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

                <AnimatedButton
                  onClick={uploadTable}
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
                      Анализ...
                    </span>
                  ) : (
                    <span className="flex items-center justify-center">
                      <AnalysisIcon />
                      <span className="ml-2">Проанализировать цены</span>
                    </span>
                  )}
                </AnimatedButton>
              </motion.div>
            </div>

            {/* Сообщение о статусе */}
            {uploadMessage && (
              <motion.div
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                className={`mt-6 rounded-xl p-4 text-center ${
                  uploadSuccess === true
                    ? "border border-green-200 bg-green-50 text-green-800"
                    : uploadSuccess === false
                      ? "border border-red-200 bg-red-50 text-red-800"
                      : "border border-blue-200 bg-blue-50 text-blue-800"
                }`}
              >
                {uploadSuccess === true && (
                  <div className="flex items-center justify-center">
                    <SuccessIcon />
                    <span className="ml-2 font-medium">{uploadMessage}</span>
                  </div>
                )}
                {uploadSuccess !== true && (
                  <span className="font-medium">{uploadMessage}</span>
                )}
              </motion.div>
            )}

            {/* Результаты анализа */}
            {showResults && analysisResults.length > 0 && (
              <motion.div
                initial={{ opacity: 0, y: 20 }}
                animate={{ opacity: 1, y: 0 }}
                transition={{ duration: 0.5, delay: 0.3 }}
                className="mt-8 rounded-2xl border border-blue-100 bg-white p-6 shadow-lg shadow-blue-100/30"
              >
                <div className="mb-6 flex items-center justify-between">
                  <div>
                    <h3 className="text-2xl font-semibold text-blue-900">
                      Результаты анализа
                    </h3>
                    <p className="mt-1 text-blue-600">
                      Найдены лучшие цены для ваших товаров
                    </p>
                  </div>
                  <span className="rounded-full bg-blue-100 px-4 py-2 text-sm font-medium text-blue-800">
                    {analysisResults.length} товаров
                  </span>
                </div>

                <div className="overflow-x-auto rounded-lg border border-blue-100">
                  <table className="w-full">
                    <thead>
                      <tr className="border-b border-blue-100 bg-blue-50">
                        <th className="px-4 py-4 text-left font-semibold text-blue-900">
                          Штрих-код
                        </th>
                        <th className="px-4 py-4 text-left font-semibold text-blue-900">
                          Наименование товара
                        </th>
                        <th className="px-4 py-4 text-left font-semibold text-blue-900">
                          Количество
                        </th>
                        <th className="px-4 py-4 text-left font-semibold text-blue-900">
                          Поставщик
                        </th>
                        <th className="px-4 py-4 text-right font-semibold text-blue-900">
                          Цена за единицу
                        </th>
                        <th className="px-4 py-4 text-center font-semibold text-blue-900">
                          Статус
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {analysisResults.map((result, index) => (
                        <motion.tr
                          key={result.barcode}
                          initial={{ opacity: 0, y: 10 }}
                          animate={{ opacity: 1, y: 0 }}
                          transition={{ delay: index * 0.05 }}
                          className="hover:bg-blue-25 border-b border-blue-50 transition-colors last:border-b-0"
                        >
                          <td className="px-4 py-4 font-mono text-sm text-blue-700">
                            {result.barcode}
                          </td>
                          <td className="px-4 py-4 text-blue-700">
                            {result.productName}
                          </td>
                          <td className="px-4 py-4 font-medium text-blue-700">
                            {result.quantity} шт
                          </td>
                          <td className="px-4 py-4 text-blue-700">
                            {result.bestSupplierName}
                          </td>
                          <td className="px-4 py-4 text-right font-semibold text-blue-900">
                            {formatPrice(result.bestPrice)} ₽
                          </td>
                          <td className="px-4 py-4 text-center">
                            <span
                              className={`inline-flex items-center rounded-full px-3 py-1 text-xs font-medium ${
                                result.requiresManualProcessing
                                  ? "border border-amber-200 bg-amber-100 text-amber-800"
                                  : "border border-green-200 bg-green-100 text-green-800"
                              }`}
                            >
                              {result.requiresManualProcessing
                                ? "Требует проверки"
                                : "Автоматически"}
                            </span>
                          </td>
                        </motion.tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                {/* Сводная информация */}
                <div className="mt-6 grid grid-cols-1 gap-4 border-t border-blue-200 pt-6 sm:grid-cols-3">
                  <div className="bg-blue-25 rounded-lg p-4 text-center">
                    <div className="text-2xl font-bold text-blue-900">
                      {analysisResults.length}
                    </div>
                    <div className="text-sm text-blue-600">Всего товаров</div>
                  </div>
                  <div className="bg-blue-25 rounded-lg p-4 text-center">
                    <div className="text-2xl font-bold text-blue-900">
                      {
                        analysisResults.filter(
                          (r) => !r.requiresManualProcessing,
                        ).length
                      }
                    </div>
                    <div className="text-sm text-blue-600">
                      Автоматически обработано
                    </div>
                  </div>
                  <div className="bg-blue-25 rounded-lg p-4 text-center">
                    <div className="text-2xl font-bold text-blue-900">
                      {formatPrice(
                        analysisResults.reduce(
                          (sum, item) => sum + item.bestPrice * item.quantity,
                          0,
                        ),
                      )}{" "}
                      ₽
                    </div>
                    <div className="text-sm text-blue-600">Общая стоимость</div>
                  </div>
                </div>
              </motion.div>
            )}

            {/* Инструкция */}
            {!showResults && (
              <motion.div
                initial={{ opacity: 0, y: 20 }}
                animate={{ opacity: 1, y: 0 }}
                transition={{ duration: 0.5, delay: 0.8 }}
                className="from-blue-25 mt-12 rounded-2xl border border-blue-100 bg-gradient-to-r to-blue-50 p-8"
              >
                <h3 className="mb-6 text-center text-2xl font-semibold text-blue-900">
                  Как работает анализ цен?
                </h3>
                <div className="grid grid-cols-1 gap-6 md:grid-cols-3">
                  <div className="p-4 text-center">
                    <div className="mx-auto mb-3 flex h-12 w-12 items-center justify-center rounded-full bg-blue-100 text-blue-600">
                      <span className="font-bold">1</span>
                    </div>
                    <h4 className="mb-2 font-semibold text-blue-900">
                      Скачайте шаблон
                    </h4>
                    <p className="text-sm text-blue-700">
                      Получите готовый файл Excel с правильной структурой для
                      заполнения
                    </p>
                  </div>
                  <div className="p-4 text-center">
                    <div className="mx-auto mb-3 flex h-12 w-12 items-center justify-center rounded-full bg-blue-100 text-blue-600">
                      <span className="font-bold">2</span>
                    </div>
                    <h4 className="mb-2 font-semibold text-blue-900">
                      Заполните данные
                    </h4>
                    <p className="text-sm text-blue-700">
                      Внесите информацию о товарах: штрих-коды и количество
                    </p>
                  </div>
                  <div className="p-4 text-center">
                    <div className="mx-auto mb-3 flex h-12 w-12 items-center justify-center rounded-full bg-blue-100 text-blue-600">
                      <span className="font-bold">3</span>
                    </div>
                    <h4 className="mb-2 font-semibold text-blue-900">
                      Получите анализ
                    </h4>
                    <p className="text-sm text-blue-700">
                      Система автоматически подберет поставщиков с лучшими
                      ценами
                    </p>
                  </div>
                </div>
              </motion.div>
            )}
          </motion.div>
        </div>
      </div>

      {/* Футер */}
      <footer className="relative z-10 border-t border-blue-100 bg-white/80 py-8 backdrop-blur-md">
        <div className="container mx-auto px-4 text-center">
          <p className="text-sm text-blue-600">
            © 2025 CleanSoul. Все права защищены.
          </p>
        </div>
      </footer>
    </div>
  );
}

export default Home;
