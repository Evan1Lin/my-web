import React, { useState } from 'react';
import { BarChart3, Lock, User, ArrowRight, Loader2, ShieldCheck } from 'lucide-react';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

interface LoginProps {
  onLogin: (token: string, username: string) => void;
}

export default function Login({ onLogin }: LoginProps) {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setError('');
    setLoading(true);

    try {
      const res = await fetch('/api/auth/login', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ username, password }),
      });

      const data = await res.json();
      if (res.ok && data.success) {
        onLogin(data.token, data.username);
      } else {
        setError(data.error || '登录失败，请检查账号和密码');
      }
    } catch (err) {
      setError('网络连接失败，请稍后再试');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="seed-login-shell flex items-center justify-center px-4 py-10 lg:px-8">
      <div className="seed-backdrop-wave" />
      <div className="seed-backdrop-grid" />
      <div className="relative z-10 w-full max-w-6xl">
        <div className="grid items-end gap-10 lg:grid-cols-[minmax(0,1fr)_440px]">
          <div className="seed-login-copy hidden lg:block">
            <div className="inline-flex items-center gap-3 rounded-full border border-white/70 bg-white/55 px-4 py-2 text-[11px] uppercase tracking-[0.22em] text-slate-500 shadow-[0_18px_45px_rgba(71,103,177,0.08)] backdrop-blur-xl">
              <span className="seed-brand-chip inline-flex size-8 items-center justify-center rounded-xl text-white">
                <BarChart3 size={16} />
              </span>
              Quality BI Dashboard
            </div>
            <h1 className="mt-8 text-6xl font-semibold tracking-[-0.07em] text-slate-950 xl:text-[88px]">
              Seed-like
              <br />
              quality cockpit
            </h1>
            <p className="mt-6 max-w-2xl text-lg leading-8 text-slate-600">
              以字节 Seed 的蓝绿流体视觉为灵感，把质量数据、返修率、OOB 和问题关闭率统一收束到一个更现代的决策界面里。
            </p>
          </div>

          <div className="seed-login-card rounded-[2rem] p-8 shadow-[0_32px_90px_rgba(71,103,177,0.16)] lg:p-9">
            <div className="flex items-center gap-3">
              <div className="seed-brand-chip flex size-12 items-center justify-center rounded-2xl text-white">
                <ShieldCheck size={22} />
              </div>
              <div>
                <p className="text-[11px] uppercase tracking-[0.22em] text-slate-500">Secure access</p>
                <h2 className="mt-1 text-2xl font-semibold tracking-[-0.05em] text-slate-900">系统登录</h2>
              </div>
            </div>

            <form onSubmit={handleSubmit} className="mt-8 space-y-5">
              <div>
                <label className="mb-1.5 block text-sm font-medium text-slate-700">用户名</label>
                <div className="relative">
                  <div className="pointer-events-none absolute inset-y-0 left-0 flex items-center pl-3">
                    <User className="text-slate-400" size={18} />
                  </div>
                  <input
                    type="text"
                    required
                    value={username}
                    onChange={(e) => setUsername(e.target.value)}
                    className="seed-input w-full rounded-2xl py-3 pl-10 pr-4 text-sm text-slate-800 transition-all focus:outline-none"
                    placeholder="请输入账号"
                  />
                </div>
              </div>

              <div>
                <label className="mb-1.5 block text-sm font-medium text-slate-700">密码</label>
                <div className="relative">
                  <div className="pointer-events-none absolute inset-y-0 left-0 flex items-center pl-3">
                    <Lock className="text-slate-400" size={18} />
                  </div>
                  <input
                    type="password"
                    required
                    value={password}
                    onChange={(e) => setPassword(e.target.value)}
                    className="seed-input w-full rounded-2xl py-3 pl-10 pr-4 text-sm text-slate-800 transition-all focus:outline-none"
                    placeholder="请输入密码"
                  />
                </div>
              </div>

              {error && (
                <div className="rounded-2xl border border-red-100 bg-red-50/85 p-3 text-sm text-red-600 animate-in fade-in zoom-in duration-200 backdrop-blur-xl">
                  {error}
                </div>
              )}

              <button
                type="submit"
                disabled={loading}
                className="seed-login-button flex w-full items-center justify-center rounded-2xl px-4 py-3 text-sm font-medium text-white transition-all hover:translate-y-[-1px] disabled:opacity-70 disabled:cursor-not-allowed"
              >
                {loading ? (
                  <Loader2 className="animate-spin" size={18} />
                ) : (
                  <>
                    登录系统
                    <ArrowRight className="ml-2" size={16} />
                  </>
                )}
              </button>
            </form>
          </div>
        </div>

        <p className="relative z-10 mt-8 text-center text-xs text-slate-400">
          &copy; {new Date().getFullYear()} Quality Intelligence. All rights reserved.
        </p>
      </div>
    </div>
  );
}
