-- İlan Çalışma Alanı — YENİ Supabase projesinde SQL Editor'de bir kez çalıştırın.
create extension if not exists pgcrypto;

create table if not exists musteri_dosyalari (
  id uuid primary key default gen_random_uuid(),
  ad text not null,
  telefon text,
  hedef text,
  butce_min numeric,
  butce_max numeric,
  bolgeler text,
  notlar text,
  durum text not null default 'acik',
  olusturma timestamptz not null default now()
);

create table if not exists ilanlar (
  id bigserial primary key,
  dosya_id uuid not null references musteri_dosyalari(id) on delete cascade,
  tur text not null check (tur in ('satilik','kiralik')),
  url text not null,
  ilan_tarihi date,
  mulk_turu text,
  ilce text,
  mahalle text,
  oda text,
  m2 numeric,
  fiyat numeric,
  ilk_fiyat numeric,
  kat text,
  toplam_kat numeric,
  bina_yasi text,
  site_icinde text,
  site_adi text,
  kullanim text,
  sahip_turu text,
  ilan_sahibi text,
  ofis text,
  yayin_suresi numeric,
  kaynak text,
  baslik text,
  -- eleme akışı: aday -> birinci (1. eleme geçti) -> ikinci (danışman görüşmesi sonrası) -> final ; elendi
  asama text not null default 'aday' check (asama in ('aday','birinci','ikinci','final','elendi')),
  elenen_asama text,
  not_metni text,
  gorusme jsonb not null default '{}'::jsonb,
  guncelleme timestamptz not null default now(),
  unique (dosya_id, tur, url)
);

create index if not exists ilanlar_dosya_asama on ilanlar (dosya_id, asama);

-- Güvenlik: tablolara yalnız uygulamanın sunucu anahtarı (service_role) erişir; anon anahtar hiçbir şey okuyamaz/yazamaz.
alter table musteri_dosyalari enable row level security;
alter table ilanlar enable row level security;
