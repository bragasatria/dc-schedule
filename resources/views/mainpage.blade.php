<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta http-equiv="X-UA-Compatible" content="ie=edge">
    <link rel="icon" type="image/png" href="{{ asset('favicon.png') }}">
    <title>DC Schedule - Engineer ATWS</title>
    @vite(['resources/css/app.css', 'resources/js/app.js'])
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
    <style>
        .drag-over {
            border-color: #0ea5e9 !important; /* sky-900 */
            background-color: #f0f9ff !important; /* sky-50 */
        }
    </style>
</head>
<body class="bg-linear-to-tr from-amber-400 to-slate-50 to-50% relative w-full overflow-y-auto min-h-screen flex flex-col items-center justify-center font-sans">

    <!-- Card Container -->
    <div class="bg-transparent backdrop-blur-none border-transparent shadow-none md:bg-white/10 md:backdrop-blur-xl md:border-white/40 md:shadow-xl p-10 rounded-2xl max-w-lg w-full transform transition duration-500 hover:scale-[1.01]">

        <!-- Header -->
        <div class="text-center mb-8">
            <div class="w-48 h-auto flex items-center justify-center mx-auto mb-8">
                <img src="https://pradita-s3.s3.ap-southeast-3.amazonaws.com/prod/alumni/logo_mitra/1693555482_ATW%20SOLAR%20PNG%20-%20HRD%20ATW%20Solar%20Indonesia.png" alt="ATW Solar">
            </div>
            <h1 class="text-3xl font-extrabold text-slate-800">DC Schedule</h1>
            <p class="text-slate-500 mt-2">Upload Excel, Dapatkan Optimasi Kabel</p>
        </div>

        <!-- Error Handling -->
        @if ($errors->any())
            <div class="bg-red-50 border-l-4 border-red-500 p-4 mb-6 rounded-r-lg animate-pulse">
                <div class="flex">
                    <div class="shrink-0">
                        <i class="fas fa-exclamation-triangle text-red-500 text-lg"></i>
                    </div>
                    <div class="ml-3">
                        <p class="text-sm text-red-700 font-semibold">Oops! Terjadi Kesalahan</p>
                        <ul class="list-disc list-inside text-xs text-red-600 mt-2 space-y-1">
                            @foreach ($errors->all() as $error)
                                <li>{{ $error }}</li>
                            @endforeach
                        </ul>
                        <p class="text-xs text-red-500 mt-3">Silakan periksa file Excel Anda dan coba lagi.</p>
                    </div>
                </div>
            </div>
        @endif

        <!-- Form -->
        <form action="{{ route('calculator.process') }}" method="POST" enctype="multipart/form-data" class="space-y-6" id="uploadForm">
            @csrf

            <!-- Upload Area -->
            <div class="relative group">
                <label for="dropzone-file" class="flex flex-col items-center justify-center w-full h-48 border-2 border-slate-300 border-dashed rounded-xl cursor-pointer bg-slate-50 hover:bg-slate-100 transition duration-300 group-hover:border-sky-900">
                    <div class="flex flex-col items-center justify-center pt-5 pb-6">
                        <i class="fa-solid fa-cloud-arrow-up text-4xl text-slate-400 mb-3 group-hover:text-sky-900 transition"></i>
                        <p class="mb-2 text-sm text-slate-600"><span class="font-semibold">Klik untuk upload</span> atau drag & drop</p>
                        <p class="text-xs text-slate-400">Format XLSX (Excel)</p>
                    </div>
                    <input id="dropzone-file" name="excel_file" type="file" class="hidden" accept=".xlsx, .xls" required />
                </label>
            </div>

            <!-- File Name Preview -->
            <div id="file-info" class="hidden bg-slate-50/40 px-4 py-3 rounded-lg items-center justify-between">
                <div class="flex items-center space-x-3">
                    <i class="fa-solid fa-file-excel text-green-800 text-lg"></i>
                    <span id="file-name" class="text-sm text-slate-700 font-medium truncate max-w-xs">filename.xlsx</span>
                </div>
                <button type="button" id="remove-file" class="text-slate-700 hover:text-red-700">
                    <i class="fa-solid fa-times"></i>
                </button>
            </div>

            <!-- Output Filename (Custom) -->
            <div class="space-y-2">
                <label for="output_name" class="ps-1 text-sm font-medium text-slate-800">Nama file output (opsional)</label>
                <input id="output_name" name="output_name" type="text" maxlength="100" placeholder="Contoh: DC Schedule - Project Name" class="w-full rounded-lg bg-slate-50/50 border border-gray-300 px-4 py-3 text-sm focus:border-sky-900 focus:outline-none focus:ring-2 focus:ring-sky-600/50" />
            </div>

            <!-- Submit Button -->
            <button type="submit" id="submitBtn" class="w-full text-white bg-linear-to-r from-sky-900  to-sky-700 hover:from-sky-900 hover:to-sky-900 focus:ring-4 focus:ring-blue-300 font-bold rounded-xl text-lg px-5 py-4 text-center shadow-lg transform active:scale-95 transition duration-200">
                <span id="btnText">Proses & Download</span>
                <span id="btnLoading" class="hidden"><i class="fas fa-spinner fa-spin mr-2"></i> Memproses...</span>
            </button>
        </form>
    </div>

    <div class="mt-8 text-slate-400 text-xs text-center mb-8">
        &copy; {{ date('Y') }} Design Engineer ATW Solar Indonesia - DC Schedule
    </div>

</body>
</html>