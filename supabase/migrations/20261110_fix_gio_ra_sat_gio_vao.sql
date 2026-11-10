-- Gio ra som hon gio vao duoi 60 phut = bam nham (lam tron 15 phut), khong phai ca qua dem -> dat gio ra = gio vao (0 gio)
create or replace function chamcong.fix_gio_ra_sat_gio_vao() returns trigger language plpgsql as $$
declare s int; e int;
begin
  if coalesce(trim(new.bat_dau),'') ~ '^\d{1,2}:\d{2}$' and coalesce(trim(new.ket_thuc),'') ~ '^\d{1,2}:\d{2}$' then
    s := split_part(trim(new.bat_dau),':',1)::int*60 + split_part(trim(new.bat_dau),':',2)::int;
    e := split_part(trim(new.ket_thuc),':',1)::int*60 + split_part(trim(new.ket_thuc),':',2)::int;
    if e < s and (s - e) < 60 then new.ket_thuc := new.bat_dau; new.gio_live := 0; end if;
  end if;
  return new;
end $$;
drop trigger if exists trg_fix_gio_ra on chamcong.ca_live;
create trigger trg_fix_gio_ra before insert or update of bat_dau, ket_thuc on chamcong.ca_live
  for each row execute function chamcong.fix_gio_ra_sat_gio_vao();
