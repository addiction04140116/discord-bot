import discord
from discord.ext import commands
from datetime import datetime
from dotenv import load_dotenv
import os
import sqlite3
import openpyxl

load_dotenv()
TOKEN = os.getenv('DISCORD_BOT_TOKEN')
bot = commands.Bot(command_prefix='!')

def init_db():
    conn = sqlite3.connect('attendance.db')
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS attendance
                 (id INTEGER PRIMARY KEY, username TEXT, start_time TEXT, end_time TEXT, daily_wage REAL)''')
    conn.commit()
    conn.close()

def record_time(username, time, type):
    conn = sqlite3.connect('attendance.db')
    c = conn.cursor()

    if type == 'start':
        c.execute("INSERT INTO attendance (username, start_time) VALUES (?, ?)", (username, time))
    elif type == 'end':
        c.execute("UPDATE attendance SET end_time = ? WHERE username = ? AND end_time IS NULL", (time, username))

    conn.commit()
    conn.close()

def calculate_wage(username, end_time):
    conn = sqlite3.connect('attendance.db')
    c = conn.cursor()
    hourly_wage = get_hourly_wage(username)

    c.execute("SELECT start_time FROM attendance WHERE username = ? AND end_time = ?", (username, end_time))
    start_time = c.fetchone()[0]
    start_time = datetime.strptime(start_time, "%Y-%m-%d %H:%M:%S.%f")
    hours_worked = (end_time - start_time).total_seconds() / 3600
    daily_wage = hours_worked * hourly_wage
    c.execute("UPDATE attendance SET daily_wage = ? WHERE username = ? AND end_time = ?", (daily_wage, username, end_time))

    conn.commit()
    conn.close()

def get_hourly_wage(username):
    return 1000  # 固定値の時給、必要に応じて変更

def export_to_excel():
    conn = sqlite3.connect('attendance.db')
    c = conn.cursor()
    c.execute("SELECT * FROM attendance")
    rows = c.fetchall()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(['ID', 'Username', 'Start Time', 'End Time', 'Daily Wage'])

    for row in rows:
        ws.append(row)

    wb.save('attendance.xlsx')
    conn.close()

@bot.event
async def on_ready():
    print(f'Logged in as {bot.user}')

@bot.command()
async def start(ctx):
    user = ctx.message.author
    start_time = datetime.now()
    await ctx.send(f'{user.mention}, 勤務開始時刻: {start_time}')
    record_time(user.name, start_time, 'start')

@bot.command()
async def end(ctx):
    user = ctx.message.author
    end_time = datetime.now()
    await ctx.send(f'{user.mention}, 勤務終了時刻: {end_time}')
    record_time(user.name, end_time, 'end')
    calculate_wage(user.name, end_time)

@bot.command()
async def export(ctx):
    export_to_excel()
    await ctx.send('データをExcelにエクスポートしました。')

# Botの起動前にデータベースを初期化
init_db()
bot.run(TOKEN)
