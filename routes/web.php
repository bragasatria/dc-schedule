<?php

use Illuminate\Support\Facades\Route;
use App\Http\Controllers\ExcelCalculatorController;

Route::get('/', [ExcelCalculatorController::class, 'index'])->name('home');
Route::post('/process', [ExcelCalculatorController::class, 'process'])->name('calculator.process');
