import math
import os
import re
import secrets
import tempfile
from datetime import datetime, date, timezone
from functools import wraps
from pathlib import Path

from dotenv import load_dotenv
from flask import (
    Flask,
    abort,
    flash,
    redirect,
    render_template,
    request,
    send_file,
    send_from_directory,
    url_for,
)
from flask_login import (
    LoginManager,
    UserMixin,
    current_user,
    login_required,
    login_user,
    logout_user,
)
from flask_migrate import Migrate
from flask_sqlalchemy import SQLAlchemy
from flask_wtf.csrf import CSRFProtect
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from sqlalchemy import UniqueConstraint, func, or_
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename

# ==============================================================================
# 1. APPLICATION & SECURITY CONFIGURATION
# ==============================================================================
load_dotenv()
BASE_DIR = Path(__file__).resolve().parent
GALLERY_DIR = BASE_DIR / "gallery"
for sub in ("albums", "players", "tournaments", "documents"):
    (GALLERY_DIR / sub).mkdir(parents=True, exist_ok=True)

app = Flask(__name__, template_folder=str(BASE_DIR / "templates"))
app.config["SECRET_KEY"] = os.getenv("SECRET_KEY", "jktta-production-secret-key-change-env")

default_sqlite_dir = Path(tempfile.gettempdir()) / "jktta_runtime_db"
default_sqlite_dir.mkdir(parents=True, exist_ok=True)
default_db_uri = f"sqlite:///{(default_sqlite_dir / 'jktta.db').as_posix()}"

db_url = os.getenv("DATABASE_URL", default_db_uri)
if db_url.startswith("postgres://"):
    db_url = db_url.replace("postgres://", "postgresql+psycopg://", 1)

app.config["SQLALCHEMY_DATABASE_URI"] = db_url
app.config["SQLALCHEMY_ENGINE_OPTIONS"] = {
    "pool_pre_ping": True,
    "pool_recycle": 280,
    "pool_size": 5,
    "max_overflow": 10,
}
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["MAX_CONTENT_LENGTH"] = 5 * 1024 * 1024  # 5 MB upload limit
app.config["GALLERY_FOLDER"] = str(GALLERY_DIR)
app.config["ALLOWED_IMAGE_EXT"] = {"png", "jpg", "jpeg", "webp"}
app.config["ALLOWED_DOC_EXT"] = {"pdf", "png", "jpg", "jpeg"}

is_production = os.getenv("FLASK_ENV") == "production"
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["SESSION_COOKIE_SECURE"] = is_production
app.config["REMEMBER_COOKIE_HTTPONLY"] = True
app.config["REMEMBER_COOKIE_SECURE"] = is_production

db = SQLAlchemy(app)
migrate = Migrate(app, db)
csrf = CSRFProtect(app)
login_manager = LoginManager(app)
login_manager.login_view = "login"
login_manager.login_message_category = "warning"


def utcnow():
    return datetime.now(timezone.utc)


def slugify_text(text: str) -> str:
    cleaned = re.sub(r"[^a-zA-Z0-9\s-]", "", text.lower()).strip()
    return re.sub(r"[\s_-]+", "-", cleaned) or secrets.token_hex(4)


def save_uploaded_file(file_storage, subfolder: str, allowed_extensions: set) -> str | None:
    if not file_storage or not file_storage.filename:
        return None
    filename = secure_filename(file_storage.filename)
    if "." not in filename:
        return None
    ext = filename.rsplit(".", 1)[1].lower()
    if ext not in allowed_extensions:
        return None
    unique_name = f"{secrets.token_hex(12)}.{ext}"
    target_dir = GALLERY_DIR / subfolder
    target_dir.mkdir(parents=True, exist_ok=True)
    file_storage.save(str(target_dir / unique_name))
    return f"{subfolder}/{unique_name}"


# ==============================================================================
# 2. COMPLETE SQLALCHEMY DATABASE MODELS
# ==============================================================================

class User(UserMixin, db.Model):
    __tablename__ = "users"

    id = db.Column(db.Integer, primary_key=True)
    email = db.Column(db.String(150), unique=True, nullable=False, index=True)
    full_name = db.Column(db.String(150), nullable=False)
    mobile = db.Column(db.String(20), nullable=True)
    password_hash = db.Column(db.String(256), nullable=False)
    role = db.Column(db.String(30), nullable=False, default="PLAYER", index=True)
    is_active = db.Column(db.Boolean, default=True, nullable=False)
    is_email_verified = db.Column(db.Boolean, default=False, nullable=False)
    verification_token = db.Column(db.String(120), nullable=True, index=True)
    reset_token = db.Column(db.String(120), nullable=True, index=True)
    reset_token_expiry = db.Column(db.DateTime(timezone=True), nullable=True)
    managed_district_id = db.Column(db.Integer, db.ForeignKey("districts.id", ondelete="SET NULL"), nullable=True)
    managed_club_id = db.Column(db.Integer, db.ForeignKey("clubs.id", ondelete="SET NULL"), nullable=True)
    last_login_at = db.Column(db.DateTime(timezone=True), nullable=True)
    created_at = db.Column(db.DateTime(timezone=True), default=utcnow, nullable=False)

    player_profile = db.relationship("Player", back_populates="user", uselist=False, cascade="all, delete-orphan")
    notifications = db.relationship("Notification", back_populates="user", cascade="all, delete-orphan")

    def set_password(self, raw_password: str) -> None:
        self.password_hash = generate_password_hash(raw_password)

    def check_password(self, raw_password: str) -> bool:
        return check_password_hash(self.password_hash, raw_password)

    @property
    def is_admin_staff(self) -> bool:
        return self.role in {
            "SUPER_ADMIN",
            "ASSOCIATION_ADMIN",
            "DISTRICT_ADMIN",
            "CLUB_ADMIN",
            "TOURNAMENT_MANAGER",
        }


class District(db.Model):
    __tablename__ = "districts"

    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), unique=True, nullable=False, index=True)
    code = db.Column(db.String(10), unique=True, nullable=False)
    division = db.Column(db.String(30), nullable=False, default="JAMMU")
    secretary_name = db.Column(db.String(150), nullable=True)
    contact_email = db.Column(db.String(150), nullable=True)
    contact_phone = db.Column(db.String(20), nullable=True)
    address = db.Column(db.Text, nullable=True)
    is_active = db.Column(db.Boolean, default=True, nullable=False)
    created_at = db.Column(db.DateTime(timezone=True), default=utcnow, nullable=False)

    clubs = db.relationship("Club", back_populates="district", cascade="all, delete-orphan")
    players = db.relationship("Player", back_populates="district")


class Club(db.Model):
    __tablename__ = "clubs"
    __table_args__ = (UniqueConstraint("name", "district_id", name="uq_club_name_district"),)

    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(150), nullable=False, index=True)
    district_id = db.Column(db.Integer, db.ForeignKey("districts.id", ondelete="CASCADE"), nullable=False, index=True)
    affiliation_number = db.Column(db.String(50), nullable=True)
    head_coach = db.Column(db.String(150), nullable=True)
    contact_email = db.Column(db.String(150), nullable=True)
    contact_phone = db.Column(db.String(20), nullable=True)
    venue_address = db.Column(db.Text, nullable=True)
    is_active = db.Column(db.Boolean, default=True, nullable=False)
    created_at = db.Column(db.DateTime(timezone=True), default=utcnow, nullable=False)

    district = db.relationship("District", back_populates="clubs")
    players = db.relationship("Player", back_populates="club")


class Player(db.Model):
    __tablename__ = "players"

    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("users.id", ondelete="CASCADE"), unique=True, nullable=False)
    registration_id = db.Column(db.String(30), unique=True, nullable=True, index=True)
    full_name = db.Column(db.String(150), nullable=False, index=True)
    date_of_birth = db.Column(db.Date, nullable=False)
    gender = db.Column(db.String(20), nullable=False, index=True)
    mobile = db.Column(db.String(20), nullable=False)
    address = db.Column(db.Text, nullable=False)
    district_id = db.Column(db.Integer, db.ForeignKey("districts.id", ondelete="RESTRICT"), nullable=False, index=True)
    club_id = db.Column(db.Integer, db.ForeignKey("clubs.id", ondelete="SET NULL"), nullable=True, index=True)
    playing_category = db.Column(db.String(50), nullable=False, index=True)
    playing_hand = db.Column(db.String(20), default="Right-Handed")
    playing_style = db.Column(db.String(50), default="Shakehand Attack")
    guardian_name = db.Column(db.String(150), nullable=True)
    guardian_mobile = db.Column(db.String(20), nullable=True)
    photo_path = db.Column(db.String(255), nullable=True)
    document_path = db.Column(db.String(255), nullable=True)
    approval_status = db.Column(db.String(30), default="PENDING", nullable=False, index=True)
    review_remarks = db.Column(db.Text, nullable=True)
    is_public_visible = db.Column(db.Boolean, default=True, nullable=False)
    is_featured = db.Column(db.Boolean, default=False, nullable=False)
    approved_at = db.Column(db.DateTime(timezone=True), nullable=True)
    created_at = db.Column(db.DateTime(timezone=True), default=utcnow, nullable=False)

    user = db.relationship("User", back_populates="player_profile")
    district = db.relationship("District", back_populates="players")
    club = db.relationship("Club", back_populates="players")
    rankings = db.relationship("PlayerRanking", back_populates="player", cascade="all, delete-orphan")
    achievements = db.relationship("Achievement", back_populates="player", cascade="all, delete-orphan")


class CommitteeMember(db.Model):
    __tablename__ = "committee_members"

    id = db.Column(db.Integer, primary_key=True)
    full_name = db.Column(db.String(150), nullable=False)
    designation = db.Column(db.String(120), nullable=False)
    district_name = db.Column(db.String(100), nullable=True)
    photo_path = db.Column(db.String(255), nullable=True)
    bio = db.Column(db.Text, nullable=True)
    display_order = db.Column(db.Integer, default=10, nullable=False)
    is_active = db.Column(db.Boolean, default=True, nullable=False)


class Tournament(db.Model):
    __tablename__ = "tournaments"

    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(200), nullable=False, index=True)
    slug = db.Column(db.String(220), unique=True, nullable=False, index=True)
    level = db.Column(db.String(30), nullable=False, default="STATE")
    district_id = db.Column(db.Integer, db.ForeignKey("districts.id", ondelete="SET NULL"), nullable=True)
    venue = db.Column(db.String(200), nullable=False)
    description = db.Column(db.Text, nullable=True)
    entry_requirements = db.Column(db.Text, nullable=True)
    banner_path = db.Column(db.String(255), nullable=True)
    start_date = db.Column(db.Date, nullable=False, index=True)
    end_date = db.Column(db.Date, nullable=False)
    registration_deadline = db.Column(db.Date, nullable=False)
    importance_multiplier = db.Column(db.Float, default=1.0, nullable=False)
    k_factor = db.Column(db.Integer, default=32, nullable=False)
    status = db.Column(db.String(30), default="UPCOMING", nullable=False, index=True)
    is_published = db.Column(db.Boolean, default=True, nullable=False)
    created_at = db.Column(db.DateTime(timezone=True), default=utcnow, nullable=False)

    categories = db.relationship("TournamentCategory", back_populates="tournament", cascade="all, delete-orphan")


class TournamentCategory(db.Model):
    __tablename__ = "tournament_categories"
    __table_args__ = (UniqueConstraint("tournament_id", "name", name="uq_tournament_category_name"),)

    id = db.Column(db.Integer, primary_key=True)
    tournament_id = db.Column(db.Integer, db.ForeignKey("tournaments.id", ondelete="CASCADE"), nullable=False, index=True)
    name = db.Column(db.String(100), nullable=False)
    gender = db.Column(db.String(20), nullable=False)
    min_age = db.Column(db.Integer, nullable=True)
    max_age = db.Column(db.Integer, nullable=True)
    format_type = db.Column(db.String(30), default="KNOCKOUT", nullable=False)
    max_participants = db.Column(db.Integer, default=64, nullable=False)
    entry_fee = db.Column(db.Numeric(10, 2), default=0, nullable=False)

    tournament = db.relationship("Tournament", back_populates="categories")
    registrations = db.relationship("TournamentRegistration", back_populates="category", cascade="all, delete-orphan")
    matches = db.relationship("Match", back_populates="category", cascade="all, delete-orphan")
    results = db.relationship("TournamentResult", back_populates="category", cascade="all, delete-orphan")


class TournamentRegistration(db.Model):
    __tablename__ = "tournament_registrations"
    __table_args__ = (
        UniqueConstraint("category_id", "player_id", name="uq_unique_player_category_entry"),
        UniqueConstraint("tournament_id", "back_number", name="uq_tournament_unique_back_number"),
    )

    id = db.Column(db.Integer, primary_key=True)
    tournament_id = db.Column(db.Integer, db.ForeignKey("tournaments.id", ondelete="CASCADE"), nullable=False, index=True)
    category_id = db.Column(db.Integer, db.ForeignKey("tournament_categories.id", ondelete="CASCADE"), nullable=False, index=True)
    player_id = db.Column(db.Integer, db.ForeignKey("players.id", ondelete="CASCADE"), nullable=False, index=True)
    back_number = db.Column(db.String(20), nullable=False, index=True)
    seed_number = db.Column(db.Integer, nullable=True)
    status = db.Column(db.String(30), default="APPROVED", nullable=False, index=True)
    registered_at = db.Column(db.DateTime(timezone=True), default=utcnow, nullable=False)

    tournament = db.relationship("Tournament")
    category = db.relationship("TournamentCategory", back_populates="registrations")
    player = db.relationship("Player")


class Match(db.Model):
    __tablename__ = "matches"

    id = db.Column(db.Integer, primary_key=True)
    tournament_id = db.Column(db.Integer, db.ForeignKey("tournaments.id", ondelete="CASCADE"), nullable=False, index=True)
    category_id = db.Column(db.Integer, db.ForeignKey("tournament_categories.id", ondelete="CASCADE"), nullable=False, index=True)
    round_number = db.Column(db.Integer, nullable=False, default=1)
    round_name = db.Column(db.String(60), nullable=False)
    match_number = db.Column(db.Integer, nullable=False, default=1)
    player1_id = db.Column(db.Integer, db.ForeignKey("players.id", ondelete="SET NULL"), nullable=True)
    player2_id = db.Column(db.Integer, db.ForeignKey("players.id", ondelete="SET NULL"), nullable=True)
    winner_id = db.Column(db.Integer, db.ForeignKey("players.id", ondelete="SET NULL"), nullable=True)
    next_match_id = db.Column(db.Integer, db.ForeignKey("matches.id", ondelete="SET NULL"), nullable=True)
    table_number = db.Column(db.String(20), nullable=True)
    scheduled_at = db.Column(db.DateTime(timezone=True), nullable=True)
    status = db.Column(db.String(30), default="SCHEDULED", nullable=False, index=True)
    is_verified = db.Column(db.Boolean, default=False, nullable=False)
    is_rating_processed = db.Column(db.Boolean, default=False, nullable=False)
    created_at = db.Column(db.DateTime(timezone=True), default=utcnow, nullable=False)

    tournament = db.relationship("Tournament")
    category = db.relationship("TournamentCategory", back_populates="matches")
    player1 = db.relationship("Player", foreign_keys=[player1_id])
    player2 = db.relationship("Player", foreign_keys=[player2_id])
    winner = db.relationship("Player", foreign_keys=[winner_id])
    scores = db.relationship("MatchScore", back_populates="match", cascade="all, delete-orphan", order_by="MatchScore.set_number")


class MatchScore(db.Model):
    __tablename__ = "match_scores"
    __table_args__ = (UniqueConstraint("match_id", "set_number", name="uq_match_set_number"),)

    id = db.Column(db.Integer, primary_key=True)
    match_id = db.Column(db.Integer, db.ForeignKey("matches.id", ondelete="CASCADE"), nullable=False, index=True)
    set_number = db.Column(db.Integer, nullable=False)
    player1_points = db.Column(db.Integer, nullable=False, default=0)
    player2_points = db.Column(db.Integer, nullable=False, default=0)

    match = db.relationship("Match", back_populates="scores")


class TournamentResult(db.Model):
    __tablename__ = "tournament_results"
    __table_args__ = (UniqueConstraint("category_id", "position", "player_id", name="uq_category_position_player"),)

    id = db.Column(db.Integer, primary_key=True)
    tournament_id = db.Column(db.Integer, db.ForeignKey("tournaments.id", ondelete="CASCADE"), nullable=False, index=True)
    category_id = db.Column(db.Integer, db.ForeignKey("tournament_categories.id", ondelete="CASCADE"), nullable=False, index=True)
    player_id = db.Column(db.Integer, db.ForeignKey("players.id", ondelete="CASCADE"), nullable=False)
    position = db.Column(db.Integer, nullable=False)
    medal = db.Column(db.String(30), nullable=True)
    points_awarded = db.Column(db.Integer, default=0, nullable=False)
    is_published = db.Column(db.Boolean, default=True, nullable=False)
    published_at = db.Column(db.DateTime(timezone=True), default=utcnow, nullable=False)

    tournament = db.relationship("Tournament")
    category = db.relationship("TournamentCategory", back_populates="results")
    player = db.relationship("Player")


class PlayerRanking(db.Model):
    __tablename__ = "player_rankings"
    __table_args__ = (UniqueConstraint("player_id", "category", name="uq_player_ranking_category"),)

    id = db.Column(db.Integer, primary_key=True)
    player_id = db.Column(db.Integer, db.ForeignKey("players.id", ondelete="CASCADE"), nullable=False, index=True)
    category = db.Column(db.String(50), nullable=False, index=True)
    elo_rating = db.Column(db.Float, default=1200.0, nullable=False, index=True)
    association_points = db.Column(db.Integer, default=0, nullable=False, index=True)
    rank_position = db.Column(db.Integer, nullable=True)
    previous_rank = db.Column(db.Integer, nullable=True)
    matches_played = db.Column(db.Integer, default=0, nullable=False)
    matches_won = db.Column(db.Integer, default=0, nullable=False)
    updated_at = db.Column(db.DateTime(timezone=True), default=utcnow, onupdate=utcnow, nullable=False)

    player = db.relationship("Player", back_populates="rankings")


class RankingHistory(db.Model):
    __tablename__ = "ranking_history"

    id = db.Column(db.Integer, primary_key=True)
    player_id = db.Column(db.Integer, db.ForeignKey("players.id", ondelete="CASCADE"), nullable=False, index=True)
    category = db.Column(db.String(50), nullable=False)
    elo_before = db.Column(db.Float, nullable=False)
    elo_after = db.Column(db.Float, nullable=False)
    points_before = db.Column(db.Integer, nullable=False)
    points_after = db.Column(db.Integer, nullable=False)
    reason = db.Column(db.String(200), nullable=False)
    recorded_at = db.Column(db.DateTime(timezone=True), default=utcnow, nullable=False)


class Achievement(db.Model):
    __tablename__ = "achievements"

    id = db.Column(db.Integer, primary_key=True)
    player_id = db.Column(db.Integer, db.ForeignKey("players.id", ondelete="CASCADE"), nullable=False, index=True)
    title = db.Column(db.String(200), nullable=False)
    event_name = db.Column(db.String(200), nullable=True)
    year = db.Column(db.Integer, nullable=False)
    medal_or_rank = db.Column(db.String(80), nullable=True)
    is_verified = db.Column(db.Boolean, default=False, nullable=False)

    player = db.relationship("Player", back_populates="achievements")


class NewsArticle(db.Model):
    __tablename__ = "news_articles"

    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(220), nullable=False)
    slug = db.Column(db.String(240), unique=True, nullable=False, index=True)
    category_name = db.Column(db.String(80), nullable=False, default="Announcement", index=True)
    summary = db.Column(db.String(400), nullable=False)
    content = db.Column(db.Text, nullable=False)
    featured_image = db.Column(db.String(255), nullable=True)
    is_published = db.Column(db.Boolean, default=True, nullable=False, index=True)
    published_at = db.Column(db.DateTime(timezone=True), default=utcnow, nullable=False, index=True)


class GalleryAlbum(db.Model):
    __tablename__ = "gallery_albums"

    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(180), nullable=False)
    slug = db.Column(db.String(200), unique=True, nullable=False)
    category = db.Column(db.String(40), nullable=False, default="TOURNAMENT")
    description = db.Column(db.Text, nullable=True)
    cover_image = db.Column(db.String(255), nullable=True)
    video_url = db.Column(db.String(300), nullable=True)
    event_date = db.Column(db.Date, default=date.today, nullable=False)
    is_published = db.Column(db.Boolean, default=True, nullable=False)

    images = db.relationship("GalleryImage", back_populates="album", cascade="all, delete-orphan")


class GalleryImage(db.Model):
    __tablename__ = "gallery_images"

    id = db.Column(db.Integer, primary_key=True)
    album_id = db.Column(db.Integer, db.ForeignKey("gallery_albums.id", ondelete="CASCADE"), nullable=False, index=True)
    image_path = db.Column(db.String(255), nullable=False)
    caption = db.Column(db.String(200), nullable=True)
    description = db.Column(db.Text, nullable=True)
    uploaded_at = db.Column(db.DateTime(timezone=True), default=utcnow, nullable=False)

    album = db.relationship("GalleryAlbum", back_populates="images")


class Notification(db.Model):
    __tablename__ = "notifications"

    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("users.id", ondelete="CASCADE"), nullable=False, index=True)
    title = db.Column(db.String(160), nullable=False)
    message = db.Column(db.Text, nullable=False)
    link_url = db.Column(db.String(255), nullable=True)
    is_read = db.Column(db.Boolean, default=False, nullable=False, index=True)
    created_at = db.Column(db.DateTime(timezone=True), default=utcnow, nullable=False)

    user = db.relationship("User", back_populates="notifications")


class AuditLog(db.Model):
    __tablename__ = "audit_logs"

    id = db.Column(db.Integer, primary_key=True)
    actor_id = db.Column(db.Integer, db.ForeignKey("users.id", ondelete="SET NULL"), nullable=True)
    action = db.Column(db.String(120), nullable=False, index=True)
    target_type = db.Column(db.String(80), nullable=True)
    target_id = db.Column(db.String(60), nullable=True)
    details = db.Column(db.Text, nullable=True)
    ip_address = db.Column(db.String(45), nullable=True)
    created_at = db.Column(db.DateTime(timezone=True), default=utcnow, nullable=False, index=True)


class SiteSetting(db.Model):
    __tablename__ = "site_settings"

    id = db.Column(db.Integer, primary_key=True)
    key = db.Column(db.String(80), unique=True, nullable=False, index=True)
    value = db.Column(db.Text, nullable=False, default="")


# ==============================================================================
# 3. JKTTA ID, TOURNAMENT BACK NUMBER, EXCEL ROSTER, FIXTURE & ELO ENGINES
# ==============================================================================

def ensure_player_jktta_id(player: Player) -> str:
    """Assigns a permanent JKTTA ID once; kept identical across all tournaments."""
    if player.registration_id:
        return player.registration_id

    district_code = player.district.code.upper() if player.district else "JK"
    year = datetime.now().year
    prefix = f"JKTTA-{district_code}-{year}-"

    existing_ids = (
        db.session.query(Player.registration_id)
        .filter(Player.registration_id.like(f"{prefix}%"))
        .all()
    )
    max_seq = 0
    for (reg_id,) in existing_ids:
        try:
            seq = int(reg_id.rsplit("-", 1)[-1])
            if seq > max_seq:
                max_seq = seq
        except (ValueError, AttributeError):
            continue

    player.registration_id = f"{prefix}{max_seq + 1:04d}"
    return player.registration_id


def generate_tournament_back_number(tournament_id: int) -> str:
    """Generates a unique Back Number per tournament (e.g. 101, 102, 103...)."""
    existing_backs = (
        db.session.query(TournamentRegistration.back_number)
        .filter_by(tournament_id=tournament_id)
        .all()
    )
    used = set()
    for (bn,) in existing_backs:
        try:
            used.add(int(bn))
        except (ValueError, TypeError):
            pass

    candidate = 101
    while candidate in used:
        candidate += 1
    return str(candidate)


def build_tournament_excel_file(tournament_id: int) -> Path:
    """Creates/updates the official .xlsx player roster for a tournament inside gallery/documents/."""
    tournament = db.session.get(Tournament, tournament_id)
    if not tournament:
        raise ValueError("Tournament not found")

    registrations = (
        TournamentRegistration.query.filter_by(tournament_id=tournament_id)
        .order_by(TournamentRegistration.id.asc())
        .all()
    )

    wb = Workbook()
    ws = wb.active
    ws.title = "Registered Players"

    ws.merge_cells("A1:K1")
    ws["A1"] = f"JKTTA OFFICIAL TOURNAMENT ROSTER — {tournament.title.upper()}"
    ws["A1"].font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill(start_color="0B1F3A", end_color="0B1F3A", fill_type="solid")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:K2")
    ws["A2"] = f"Venue: {tournament.venue} | Dates: {tournament.start_date} to {tournament.end_date}"
    ws["A2"].font = Font(name="Calibri", size=10, italic=True, color="1A2332")
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 20

    headers = [
        "S.No",
        "Back No.",
        "JKTTA ID",
        "Player Full Name",
        "Gender",
        "Date of Birth",
        "Category",
        "District",
        "Club",
        "Status",
        "Registered On",
    ]
    header_fill = PatternFill(start_color="D4AF37", end_color="D4AF37", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="071426")

    ws.append(headers)
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=3, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[3].height = 24

    for idx, reg in enumerate(registrations, start=1):
        p = reg.player
        ensure_player_jktta_id(p)
        ws.append([
            idx,
            reg.back_number,
            p.registration_id,
            p.full_name,
            p.gender,
            p.date_of_birth.strftime("%Y-%m-%d") if p.date_of_birth else "",
            reg.category.name if reg.category else p.playing_category,
            p.district.name if p.district else "",
            p.club.name if p.club else "Independent",
            reg.status,
            reg.registered_at.strftime("%Y-%m-%d %H:%M") if reg.registered_at else "",
        ])

    for col_idx, width in enumerate([8, 12, 22, 26, 12, 15, 24, 18, 22, 15, 20], start=1):
        ws.column_dimensions[ws.cell(row=3, column=col_idx).column_letter].width = width

    output_path = GALLERY_DIR / "documents" / f"tournament_{tournament.id}_players.xlsx"
    wb.save(str(output_path))
    return output_path


def register_player_for_tournament(player: Player, category: TournamentCategory) -> tuple[TournamentRegistration, Path]:
    existing = TournamentRegistration.query.filter_by(category_id=category.id, player_id=player.id).first()
    if existing:
        raise ValueError("You are already registered in this tournament category.")

    same_tournament_entry = TournamentRegistration.query.filter_by(
        tournament_id=category.tournament_id,
        player_id=player.id,
    ).first()

    ensure_player_jktta_id(player)
    back_no = (
        same_tournament_entry.back_number
        if same_tournament_entry
        else generate_tournament_back_number(category.tournament_id)
    )

    registration = TournamentRegistration(
        tournament_id=category.tournament_id,
        category_id=category.id,
        player_id=player.id,
        back_number=back_no,
        status="APPROVED",
    )
    db.session.add(registration)
    db.session.commit()

    excel_path = build_tournament_excel_file(category.tournament_id)
    return registration, excel_path


def generate_knockout_fixtures(category: TournamentCategory) -> int:
    """Generates Round 1 knockout fixtures for approved entries without duplicates."""
    if Match.query.filter_by(category_id=category.id).count() > 0:
        raise ValueError("Fixtures have already been generated for this category.")

    entries = (
        TournamentRegistration.query.filter_by(category_id=category.id, status="APPROVED")
        .order_by(TournamentRegistration.seed_number.asc().nullslast(), TournamentRegistration.id.asc())
        .all()
    )
    if len(entries) < 2:
        raise ValueError("At least 2 approved players are required to generate fixtures.")

    bracket_size = 2 ** math.ceil(math.log2(len(entries)))
    round_label = {2: "Final", 4: "Semi-Final", 8: "Quarter-Final", 16: "Round of 16", 32: "Round of 32"}.get(
        bracket_size, f"Round of {bracket_size}"
    )

    padded_players = [e.player_id for e in entries] + [None] * (bracket_size - len(entries))
    created_count = 0
    for idx in range(0, bracket_size, 2):
        p1_id = padded_players[idx]
        p2_id = padded_players[idx + 1]
        is_bye = p1_id is not None and p2_id is None
        match = Match(
            tournament_id=category.tournament_id,
            category_id=category.id,
            round_number=1,
            round_name=round_label,
            match_number=(idx // 2) + 1,
            player1_id=p1_id,
            player2_id=p2_id,
            winner_id=p1_id if is_bye else None,
            status="WALKOVER" if is_bye else "SCHEDULED",
            is_verified=is_bye,
        )
        db.session.add(match)
        created_count += 1

    db.session.commit()
    return created_count


def apply_verified_match_ratings(match: Match) -> None:
    """Updates Elo rating and Association Points once after a verified match result."""
    if not match.is_verified or match.is_rating_processed or not match.player1_id or not match.player2_id or not match.winner_id:
        return

    cat_name = match.player1.playing_category or "Open"
    k_factor = match.tournament.k_factor if match.tournament else 32
    mult = match.tournament.importance_multiplier if match.tournament else 1.0

    r1 = PlayerRanking.query.filter_by(player_id=match.player1_id, category=cat_name).first()
    if not r1:
        r1 = PlayerRanking(player_id=match.player1_id, category=cat_name, elo_rating=1200.0, association_points=0)
        db.session.add(r1)

    r2 = PlayerRanking.query.filter_by(player_id=match.player2_id, category=cat_name).first()
    if not r2:
        r2 = PlayerRanking(player_id=match.player2_id, category=cat_name, elo_rating=1200.0, association_points=0)
        db.session.add(r2)

    db.session.flush()

    e1 = 1.0 / (1.0 + 10.0 ** ((r2.elo_rating - r1.elo_rating) / 400.0))
    e2 = 1.0 - e1
    s1 = 1.0 if match.winner_id == match.player1_id else 0.0
    s2 = 1.0 - s1

    elo1_before, elo2_before = r1.elo_rating, r2.elo_rating
    pts1_before, pts2_before = r1.association_points, r2.association_points

    r1.elo_rating = round(elo1_before + k_factor * (s1 - e1), 2)
    r2.elo_rating = round(elo2_before + k_factor * (s2 - e2), 2)
    r1.matches_played += 1
    r2.matches_played += 1

    if s1 == 1.0:
        r1.matches_won += 1
        r1.association_points += int(round(25 * mult))
        r2.association_points += int(round(5 * mult))
    else:
        r2.matches_won += 1
        r2.association_points += int(round(25 * mult))
        r1.association_points += int(round(5 * mult))

    db.session.add(
        RankingHistory(
            player_id=r1.player_id,
            category=cat_name,
            elo_before=elo1_before,
            elo_after=r1.elo_rating,
            points_before=pts1_before,
            points_after=r1.association_points,
            reason=f"Verified Match #{match.id} ({match.round_name})",
        )
    )
    db.session.add(
        RankingHistory(
            player_id=r2.player_id,
            category=cat_name,
            elo_before=elo2_before,
            elo_after=r2.elo_rating,
            points_before=pts2_before,
            points_after=r2.association_points,
            reason=f"Verified Match #{match.id} ({match.round_name})",
        )
    )

    # Recompute rank positions within category
    ranked_rows = (
        PlayerRanking.query.filter_by(category=cat_name)
        .order_by(PlayerRanking.elo_rating.desc(), PlayerRanking.association_points.desc())
        .all()
    )
    for pos, row in enumerate(ranked_rows, start=1):
        row.previous_rank = row.rank_position
        row.rank_position = pos

    match.is_rating_processed = True


# ==============================================================================
# 4. AUTHENTICATION, RBAC & TEMPLATE CONTEXT
# ==============================================================================

@login_manager.user_loader
def load_user(user_id):
    return db.session.get(User, int(user_id))


def admin_required(fn):
    @wraps(fn)
    @login_required
    def wrapper(*args, **kwargs):
        if not current_user.is_admin_staff:
            abort(403)
        return fn(*args, **kwargs)
    return wrapper


def log_audit(action: str, target_type: str = "", target_id: str = "", details: str = "") -> None:
    db.session.add(
        AuditLog(
            actor_id=current_user.id if current_user.is_authenticated else None,
            action=action,
            target_type=target_type,
            target_id=str(target_id),
            details=details,
            ip_address=request.remote_addr if request else None,
        )
    )


def get_setting(key: str, default: str = "") -> str:
    row = SiteSetting.query.filter_by(key=key).first()
    return row.value if row and row.value else default


def set_setting(key: str, value: str) -> None:
    row = SiteSetting.query.filter_by(key=key).first()
    if row:
        row.value = value
    else:
        db.session.add(SiteSetting(key=key, value=value))


@app.context_processor
def inject_globals() :
    return {
        "jktta": {
            "ORG_NAME": get_setting("org_name", "Jammu & Kashmir Table Tennis Association"),
            "SHORT_NAME": get_setting("short_name", "JKTTA"),
            "CONTACT_EMAIL": get_setting("contact_email", "jktta60@gmail.com"),
            "CONTACT_PHONE": get_setting("contact_phone", "9419194654"),
            "OFFICE_ADDRESS": get_setting(
                "office_address",
                "TTAJK Office: 112, Raghunath Pura, Jammu, Jammu & Kashmir - 180001",
            ),
            "PRESIDENT_NAME": get_setting("president_name", "Sanjay Wahi"),
            "SECRETARY_NAME": get_setting("secretary_name", "Rajan Sharma"),
            "TREASURER_NAME": get_setting("treasurer_name", "Narinder Singh"),
            "PRESIDENT_MESSAGE": get_setting(
                "president_message",
                "Welcome to the 49th UTT J&K Table Tennis UT Championship – 2026 at Moulana Azad Stadium, Jammu.",
            ),
            "SECRETARY_MESSAGE": get_setting(
                "secretary_message",
                "All entries for the 49th UTT J&K Table Tennis UT Championship – 2026 close on 7th December 2026 at 5:00 PM.",
            ),
            "ABOUT_HISTORY": get_setting(
                "about_history",
                "Conducting state and district championships across Jammu & Kashmir in collaboration with the Jammu & Kashmir Sports Council under the aegis of the Table Tennis Federation of India (TTFI).",
            ),
            "VISION": get_setting(
                "vision",
                "To nurture disciplined, nationally competitive table tennis athletes from every district of Jammu & Kashmir.",
            ),
            "MISSION": get_setting(
                "mission",
                "To organize transparent TTFI-affiliated UT and District ranking championships and develop grassroots table tennis across J&K.",
            ),
            "MAP_EMBED_URL": get_setting("map_embed_url", ""),
        },
        "current_year": datetime.now().year,
    }


# ==============================================================================
# 5. MEDIA, PUBLIC PAGES & EXCEL EXPORT ROUTES
# ==============================================================================

@app.route("/gallery/<path:filename>")
def serve_gallery(filename):
    if filename.startswith("documents/") and filename != "documents/PROSPECTUS_JKTTA_UT_TOURNAMENT_2026-27.pdf":
        if not current_user.is_authenticated:
            abort(403)
    return send_from_directory(app.config["GALLERY_FOLDER"], filename)

@app.route("/prospectus")
def official_prospectus():
    pdf_path = GALLERY_DIR / "documents" / "PROSPECTUS_JKTTA_UT_TOURNAMENT_2026-27.pdf"
    if not pdf_path.exists():
        abort(404)
    return send_file(
        str(pdf_path),
        mimetype="application/pdf",
        as_attachment=False,
        download_name="PROSPECTUS_JKTTA_UT_TOURNAMENT_2026-27.pdf",
    )


@app.route("/tournaments/<int:tournament_id>/export-excel")
@login_required
def download_tournament_excel(tournament_id):
    tournament = db.session.get(Tournament, tournament_id)
    if not tournament:
        abort(404)

    if not current_user.is_admin_staff:
        player = current_user.player_profile
        is_registered = (
            TournamentRegistration.query.filter_by(tournament_id=tournament_id, player_id=player.id).first()
            if player
            else None
        )
        if not is_registered:
            abort(403)

    excel_path = build_tournament_excel_file(tournament_id)
    return send_file(
        str(excel_path),
        as_attachment=True,
        download_name=f"JKTTA_{slugify_text(tournament.title)}_Roster.xlsx",
    )


@app.route("/")
def home():
    stats = {
        "approved_players": Player.query.filter_by(approval_status="APPROVED").count(),
        "districts_count": District.query.filter_by(is_active=True).count(),
        "clubs_count": Club.query.filter_by(is_active=True).count(),
        "tournaments_count": Tournament.query.filter_by(is_published=True).count(),
    }
    upcoming_tournaments = (
        Tournament.query.filter(Tournament.is_published.is_(True), Tournament.status.in_(["UPCOMING", "ONGOING"]))
        .order_by(Tournament.start_date.asc())
        .limit(4)
        .all()
    )
    featured_players = (
        Player.query.filter_by(approval_status="APPROVED", is_public_visible=True, is_featured=True)
        .limit(4)
        .all()
    )
    top_rankings = (
        PlayerRanking.query.join(Player)
        .filter(Player.approval_status == "APPROVED", Player.is_public_visible.is_(True))
        .order_by(PlayerRanking.elo_rating.desc())
        .limit(6)
        .all()
    )
    latest_results = (
        TournamentResult.query.filter_by(is_published=True)
        .order_by(TournamentResult.published_at.desc())
        .limit(5)
        .all()
    )
    latest_news = (
        NewsArticle.query.filter_by(is_published=True)
        .order_by(NewsArticle.published_at.desc())
        .limit(3)
        .all()
    )
    gallery_albums = (
        GalleryAlbum.query.filter_by(is_published=True)
        .order_by(GalleryAlbum.event_date.desc())
        .limit(4)
        .all()
    )
    return render_template(
        "home.html",
        stats=stats,
        upcoming_tournaments=upcoming_tournaments,
        featured_players=featured_players,
        top_rankings=top_rankings,
        latest_results=latest_results,
        latest_news=latest_news,
        gallery_albums=gallery_albums,
    )


@app.route("/about")
def about():
    committee = (
        CommitteeMember.query.filter_by(is_active=True)
        .order_by(CommitteeMember.display_order.asc(), CommitteeMember.full_name.asc())
        .all()
    )
    districts = District.query.filter_by(is_active=True).order_by(District.name.asc()).all()
    clubs = Club.query.filter_by(is_active=True).order_by(Club.name.asc()).all()
    return render_template("about.html", committee=committee, districts=districts, clubs=clubs)


@app.route("/players")
def players_directory():
    q = request.args.get("q", "").strip()
    district_id = request.args.get("district", type=int)
    club_id = request.args.get("club", type=int)
    gender = request.args.get("gender", "").strip()
    category = request.args.get("category", "").strip()

    query = Player.query.filter_by(approval_status="APPROVED", is_public_visible=True)
    if q:
        like_q = f"%{q}%"
        query = query.filter(or_(Player.full_name.ilike(like_q), Player.registration_id.ilike(like_q)))
    if district_id:
        query = query.filter(Player.district_id == district_id)
    if club_id:
        query = query.filter(Player.club_id == club_id)
    if gender:
        query = query.filter(Player.gender == gender)
    if category:
        query = query.filter(Player.playing_category == category)

    players = query.order_by(Player.full_name.asc()).all()
    districts = District.query.filter_by(is_active=True).order_by(District.name.asc()).all()
    clubs = Club.query.filter_by(is_active=True).order_by(Club.name.asc()).all()
    categories = ["U-11", "U-13", "U-15", "U-17", "U-19", "Men", "Women", "Veteran"]

    return render_template(
        "players.html",
        players=players,
        districts=districts,
        clubs=clubs,
        categories=categories,
        filters={"q": q, "district": district_id, "club": club_id, "gender": gender, "category": category},
    )


@app.route("/players/<int:player_id>")
def player_detail(player_id):
    player = Player.query.filter_by(id=player_id, approval_status="APPROVED", is_public_visible=True).first_or_404()
    rankings = PlayerRanking.query.filter_by(player_id=player.id).all()
    achievements = Achievement.query.filter_by(player_id=player.id, is_verified=True).order_by(Achievement.year.desc()).all()
    recent_matches = (
        Match.query.filter(
            Match.is_verified.is_(True),
            or_(Match.player1_id == player.id, Match.player2_id == player.id),
        )
        .order_by(Match.id.desc())
        .limit(10)
        .all()
    )
    tournament_entries = (
        TournamentRegistration.query.filter_by(player_id=player.id)
        .order_by(TournamentRegistration.registered_at.desc())
        .all()
    )
    return render_template(
        "player_detail.html",
        player=player,
        rankings=rankings,
        achievements=achievements,
        recent_matches=recent_matches,
        tournament_entries=tournament_entries,
    )


@app.route("/tournaments")
def tournaments_list():
    status_filter = request.args.get("status", "").strip().upper()
    level_filter = request.args.get("level", "").strip().upper()

    query = Tournament.query.filter_by(is_published=True)
    if status_filter in {"UPCOMING", "ONGOING", "COMPLETED"}:
        query = query.filter(Tournament.status == status_filter)
    if level_filter in {"DISTRICT", "STATE", "OPEN"}:
        query = query.filter(Tournament.level == level_filter)

    tournaments = query.order_by(Tournament.start_date.desc()).all()
    return render_template(
        "tournaments.html",
        tournaments=tournaments,
        status_filter=status_filter,
        level_filter=level_filter,
    )


@app.route("/tournaments/<slug>", methods=["GET", "POST"])
def tournament_detail(slug):
    tournament = Tournament.query.filter_by(slug=slug, is_published=True).first_or_404()

    if request.method == "POST":
        if not current_user.is_authenticated:
            flash("Please log in to register for this tournament.", "warning")
            return redirect(url_for("login", next=url_for("tournament_detail", slug=slug)))

        player = current_user.player_profile
        if not player or player.approval_status != "APPROVED":
            flash("Only approved JKTTA players can register for tournaments.", "danger")
            return redirect(url_for("tournament_detail", slug=slug))

        if date.today() > tournament.registration_deadline or tournament.status not in {"UPCOMING", "ONGOING"}:
            flash("Registration deadline for this tournament has closed.", "danger")
            return redirect(url_for("tournament_detail", slug=slug))

        category_id = request.form.get("category_id", type=int)
        category = TournamentCategory.query.filter_by(id=category_id, tournament_id=tournament.id).first()
        if not category:
            flash("Please select a valid tournament category.", "danger")
            return redirect(url_for("tournament_detail", slug=slug))

        try:
            reg, _ = register_player_for_tournament(player, category)
            log_audit(
                "TOURNAMENT_REGISTER",
                "TournamentRegistration",
                str(reg.id),
                f"{player.registration_id} (Back #{reg.back_number}) registered in {tournament.title} - {category.name}",
            )
            db.session.commit()
            flash(
                f"Registered successfully! Permanent JKTTA ID: {player.registration_id} | "
                f"Tournament Back No.: #{reg.back_number}. Excel roster updated.",
                "success",
            )
        except ValueError as exc:
            flash(str(exc), "warning")

        return redirect(url_for("tournament_detail", slug=slug))

    registrations = (
        TournamentRegistration.query.filter_by(tournament_id=tournament.id)
        .order_by(TournamentRegistration.back_number.asc())
        .all()
    )
    matches = (
        Match.query.filter_by(tournament_id=tournament.id)
        .order_by(Match.round_number.asc(), Match.match_number.asc())
        .all()
    )
    user_registrations = []
    if current_user.is_authenticated and current_user.player_profile:
        user_registrations = TournamentRegistration.query.filter_by(
            tournament_id=tournament.id,
            player_id=current_user.player_profile.id,
        ).all()

    return render_template(
        "tournament_detail.html",
        tournament=tournament,
        registrations=registrations,
        matches=matches,
        user_registrations=user_registrations,
    )


@app.route("/rankings")
def rankings_page():
    category = request.args.get("category", "").strip()
    district_id = request.args.get("district", type=int)
    sort_by = request.args.get("sort", "elo").strip().lower()
    q = request.args.get("q", "").strip()

    query = (
        PlayerRanking.query.join(Player)
        .filter(Player.approval_status == "APPROVED", Player.is_public_visible.is_(True))
    )
    if category:
        query = query.filter(PlayerRanking.category == category)
    if district_id:
        query = query.filter(Player.district_id == district_id)
    if q:
        like_q = f"%{q}%"
        query = query.filter(or_(Player.full_name.ilike(like_q), Player.registration_id.ilike(like_q)))

    if sort_by == "points":
        query = query.order_by(PlayerRanking.association_points.desc(), PlayerRanking.elo_rating.desc())
    else:
        query = query.order_by(PlayerRanking.elo_rating.desc(), PlayerRanking.association_points.desc())

    rankings = query.all()
    districts = District.query.filter_by(is_active=True).order_by(District.name.asc()).all()
    categories = ["U-11", "U-13", "U-15", "U-17", "U-19", "Men", "Women", "Veteran"]

    return render_template(
        "rankings.html",
        rankings=rankings,
        districts=districts,
        categories=categories,
        filters={"category": category, "district": district_id, "sort": sort_by, "q": q},
    )


@app.route("/results")
def results_page():
    tournament_id = request.args.get("tournament", type=int)
    query = TournamentResult.query.filter_by(is_published=True)
    if tournament_id:
        query = query.filter(TournamentResult.tournament_id == tournament_id)

    results = query.order_by(TournamentResult.published_at.desc(), TournamentResult.position.asc()).all()
    tournaments = Tournament.query.filter_by(is_published=True).order_by(Tournament.start_date.desc()).all()
    recent_matches = (
        Match.query.filter_by(is_verified=True, status="COMPLETED")
        .order_by(Match.id.desc())
        .limit(20)
        .all()
    )
    return render_template(
        "results.html",
        results=results,
        tournaments=tournaments,
        recent_matches=recent_matches,
        selected_tournament=tournament_id,
    )


@app.route("/news")
def news_list():
    q = request.args.get("q", "").strip()
    category = request.args.get("category", "").strip()

    query = NewsArticle.query.filter_by(is_published=True)
    if q:
        like_q = f"%{q}%"
        query = query.filter(or_(NewsArticle.title.ilike(like_q), NewsArticle.summary.ilike(like_q)))
    if category:
        query = query.filter(NewsArticle.category_name == category)

    articles = query.order_by(NewsArticle.published_at.desc()).all()
    categories = [row[0] for row in db.session.query(NewsArticle.category_name).distinct().all()]
    return render_template("news.html", articles=articles, categories=categories, filters={"q": q, "category": category})


@app.route("/news/<slug>")
def news_detail(slug):
    article = NewsArticle.query.filter_by(slug=slug, is_published=True).first_or_404()
    related = (
        NewsArticle.query.filter(NewsArticle.is_published.is_(True), NewsArticle.id != article.id)
        .order_by(NewsArticle.published_at.desc())
        .limit(3)
        .all()
    )
    return render_template("news_detail.html", article=article, related=related)


@app.route("/gallery")
def gallery_page():
    category = request.args.get("category", "").strip().upper()
    query = GalleryAlbum.query.filter_by(is_published=True)
    if category in {"TOURNAMENT", "TRAINING_CAMP", "AWARD_CEREMONY", "ASSOCIATION_EVENT"}:
        query = query.filter(GalleryAlbum.category == category)
    albums = query.order_by(GalleryAlbum.event_date.desc()).all()
    return render_template("gallery.html", albums=albums, selected_category=category)


@app.route("/contact", methods=["GET", "POST"])
def contact_page():
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        email = request.form.get("email", "").strip()
        subject = request.form.get("subject", "").strip()
        message = request.form.get("message", "").strip()

        if not name or not email or not subject or not message:
            flash("Please complete all required fields in the contact form.", "danger")
        else:
            log_audit("CONTACT_INQUIRY", "Contact", email, f"From {name} ({email}) | Subject: {subject} | {message}")
            db.session.commit()
            flash("Thank you! Your inquiry has been logged with the JKTTA Secretariat.", "success")
            return redirect(url_for("contact_page"))

    districts = District.query.filter_by(is_active=True).order_by(District.name.asc()).all()
    return render_template("contact.html", districts=districts)


# ==============================================================================
# 6. AUTHENTICATION, REGISTRATION & PLAYER PORTAL ROUTES
# ==============================================================================

@app.route("/login", methods=["GET", "POST"])
def login():
    if current_user.is_authenticated:
        return redirect(url_for("admin_dashboard") if current_user.is_admin_staff else url_for("player_dashboard"))

    if request.method == "POST":
        email = request.form.get("email", "").strip().lower()
        password = request.form.get("password", "")
        remember = bool(request.form.get("remember"))

        user = User.query.filter(func.lower(User.email) == email).first()
        if user and user.check_password(password):
            if not user.is_active:
                flash("Your account is inactive or suspended. Contact JKTTA Secretariat.", "danger")
                return render_template("auth/login.html")
            login_user(user, remember=remember)
            user.last_login_at = utcnow()
            log_audit("USER_LOGIN", "User", str(user.id), f"Login: {user.email}")
            db.session.commit()
            flash(f"Welcome back, {user.full_name}!", "success")
            default_dest = url_for("admin_dashboard") if user.is_admin_staff else url_for("player_dashboard")
            return redirect(request.args.get("next") or default_dest)

        flash("Invalid email address or password.", "danger")

    return render_template("auth/login.html")


@app.route("/logout", methods=["POST"])
@login_required
def logout():
    log_audit("USER_LOGOUT", "User", str(current_user.id), f"Logout: {current_user.email}")
    db.session.commit()
    logout_user()
    flash("Signed out securely.", "info")
    return redirect(url_for("home"))


@app.route("/register", methods=["GET", "POST"])
def register():
    if current_user.is_authenticated:
        return redirect(url_for("player_dashboard"))

    districts = District.query.filter_by(is_active=True).order_by(District.name.asc()).all()
    clubs = Club.query.filter_by(is_active=True).order_by(Club.name.asc()).all()
    categories = ["U-11", "U-13", "U-15", "U-17", "U-19", "Men", "Women", "Veteran"]

    if request.method == "POST":
        full_name = request.form.get("full_name", "").strip()
        email = request.form.get("email", "").strip().lower()
        mobile = request.form.get("mobile", "").strip()
        password = request.form.get("password", "")
        confirm_password = request.form.get("confirm_password", "")
        dob_str = request.form.get("date_of_birth", "").strip()
        gender = request.form.get("gender", "").strip()
        address = request.form.get("address", "").strip()
        district_id = request.form.get("district_id", type=int)
        club_id = request.form.get("club_id", type=int) or None
        playing_category = request.form.get("playing_category", "").strip()
        playing_hand = request.form.get("playing_hand", "Right-Handed").strip()
        playing_style = request.form.get("playing_style", "Shakehand Attack").strip()
        guardian_name = request.form.get("guardian_name", "").strip()
        guardian_mobile = request.form.get("guardian_mobile", "").strip()

        if not all([full_name, email, mobile, password, dob_str, gender, address, district_id, playing_category]):
            flash("Please fill in all required player registration fields.", "danger")
            return render_template("auth/register.html", districts=districts, clubs=clubs, categories=categories)

        if password != confirm_password or len(password) < 8:
            flash("Passwords must match and be at least 8 characters long.", "danger")
            return render_template("auth/register.html", districts=districts, clubs=clubs, categories=categories)

        if User.query.filter(func.lower(User.email) == email).first():
            flash("An account with this email already exists. Please log in instead.", "warning")
            return redirect(url_for("login"))

        try:
            dob = datetime.strptime(dob_str, "%Y-%m-%d").date()
        except ValueError:
            flash("Invalid Date of Birth format.", "danger")
            return render_template("auth/register.html", districts=districts, clubs=clubs, categories=categories)

        photo_file = request.files.get("photo")
        doc_file = request.files.get("document")
        photo_path = save_uploaded_file(photo_file, "players", app.config["ALLOWED_IMAGE_EXT"]) if photo_file and photo_file.filename else None
        doc_path = save_uploaded_file(doc_file, "documents", app.config["ALLOWED_DOC_EXT"]) if doc_file and doc_file.filename else None

        verify_token = secrets.token_urlsafe(32)
        user = User(
            email=email,
            full_name=full_name,
            mobile=mobile,
            role="PLAYER",
            is_active=True,
            is_email_verified=False,
            verification_token=verify_token,
        )
        user.set_password(password)
        db.session.add(user)
        db.session.flush()

        player = Player(
            user_id=user.id,
            full_name=full_name,
            date_of_birth=dob,
            gender=gender,
            mobile=mobile,
            address=address,
            district_id=district_id,
            club_id=club_id,
            playing_category=playing_category,
            playing_hand=playing_hand,
            playing_style=playing_style,
            guardian_name=guardian_name,
            guardian_mobile=guardian_mobile,
            photo_path=photo_path,
            document_path=doc_path,
            approval_status="PENDING",
        )
        db.session.add(player)
        db.session.add(
            Notification(
                user_id=user.id,
                title="Registration Application Submitted",
                message="Your JKTTA player profile is pending District/Association verification.",
                link_url="/player/dashboard",
            )
        )
        log_audit("PLAYER_REGISTERED", "Player", str(player.id), f"New player application: {full_name} ({email})")
        db.session.commit()

        flash("Registration submitted! Your profile is awaiting admin verification.", "success")
        login_user(user)
        return redirect(url_for("player_dashboard"))

    return render_template("auth/register.html", districts=districts, clubs=clubs, categories=categories)


@app.route("/verify-email/<token>")
def verify_email(token):
    user = User.query.filter_by(verification_token=token).first()
    if not user:
        flash("Invalid or expired verification link.", "danger")
        return redirect(url_for("home"))
    user.is_email_verified = True
    user.verification_token = None
    db.session.commit()
    flash("Email verified successfully!", "success")
    return redirect(url_for("player_dashboard") if current_user.is_authenticated else url_for("login"))


@app.route("/forgot-password", methods=["GET", "POST"])
def forgot_password():
    if request.method == "POST":
        email = request.form.get("email", "").strip().lower()
        user = User.query.filter(func.lower(User.email) == email).first()
        if user:
            token = secrets.token_urlsafe(32)
            user.reset_token = token
            user.reset_token_expiry = utcnow()
            db.session.commit()
            print(f"[JKTTA PASSWORD RESET] {email} -> {url_for('reset_password', token=token, _external=True)}")
        flash("If that email is registered, a password reset link has been generated.", "info")
        return redirect(url_for("login"))
    return render_template("auth/forgot_password.html")


@app.route("/reset-password/<token>", methods=["GET", "POST"])
def reset_password(token):
    user = User.query.filter_by(reset_token=token).first()
    if not user:
        flash("Invalid or expired password reset token.", "danger")
        return redirect(url_for("forgot_password"))

    if request.method == "POST":
        password = request.form.get("password", "")
        confirm = request.form.get("confirm_password", "")
        if password != confirm or len(password) < 8:
            flash("Passwords must match and be at least 8 characters.", "danger")
            return render_template("auth/reset_password.html", token=token)
        user.set_password(password)
        user.reset_token = None
        user.reset_token_expiry = None
        db.session.commit()
        flash("Password updated successfully. Please sign in.", "success")
        return redirect(url_for("login"))

    return render_template("auth/reset_password.html", token=token)


@app.route("/player/dashboard")
@login_required
def player_dashboard():
    if current_user.is_admin_staff and not current_user.player_profile:
        return redirect(url_for("admin_dashboard"))

    player = current_user.player_profile
    if not player:
        flash("No player profile is linked to this account.", "info")
        return redirect(url_for("home"))

    rankings = PlayerRanking.query.filter_by(player_id=player.id).all()
    registrations = (
        TournamentRegistration.query.filter_by(player_id=player.id)
        .order_by(TournamentRegistration.registered_at.desc())
        .all()
    )
    upcoming_tournaments = (
        Tournament.query.filter(Tournament.is_published.is_(True), Tournament.status.in_(["UPCOMING", "ONGOING"]))
        .order_by(Tournament.start_date.asc())
        .limit(5)
        .all()
    )
    matches = (
        Match.query.filter(or_(Match.player1_id == player.id, Match.player2_id == player.id))
        .order_by(Match.id.desc())
        .limit(10)
        .all()
    )
    results = TournamentResult.query.filter_by(player_id=player.id).order_by(TournamentResult.published_at.desc()).all()
    achievements = Achievement.query.filter_by(player_id=player.id).order_by(Achievement.year.desc()).all()
    notifications = (
        Notification.query.filter_by(user_id=current_user.id)
        .order_by(Notification.created_at.desc())
        .limit(10)
        .all()
    )

    fields = [
        player.full_name, player.date_of_birth, player.gender, player.mobile,
        player.address, player.district_id, player.playing_category, player.photo_path, player.document_path,
    ]
    completion_pct = int((sum(1 for f in fields if f) / len(fields)) * 100)

    return render_template(
        "player/dashboard.html",
        player=player,
        rankings=rankings,
        registrations=registrations,
        upcoming_tournaments=upcoming_tournaments,
        matches=matches,
        results=results,
        achievements=achievements,
        notifications=notifications,
        completion_pct=completion_pct,
    )


@app.route("/player/edit-profile", methods=["GET", "POST"])
@login_required
def edit_player_profile():
    player = current_user.player_profile
    if not player:
        abort(404)

    districts = District.query.filter_by(is_active=True).order_by(District.name.asc()).all()
    clubs = Club.query.filter_by(is_active=True).order_by(Club.name.asc()).all()
    categories = ["U-11", "U-13", "U-15", "U-17", "U-19", "Men", "Women", "Veteran"]

    if request.method == "POST":
        player.full_name = request.form.get("full_name", player.full_name).strip()
        current_user.full_name = player.full_name
        player.mobile = request.form.get("mobile", player.mobile).strip()
        current_user.mobile = player.mobile
        player.address = request.form.get("address", player.address).strip()
        player.club_id = request.form.get("club_id", type=int) or None
        player.playing_category = request.form.get("playing_category", player.playing_category).strip()

        photo_file = request.files.get("photo")
        if photo_file and photo_file.filename:
            saved_photo = save_uploaded_file(photo_file, "players", app.config["ALLOWED_IMAGE_EXT"])
            if saved_photo:
                player.photo_path = saved_photo

        doc_file = request.files.get("document")
        if doc_file and doc_file.filename:
            saved_doc = save_uploaded_file(doc_file, "documents", app.config["ALLOWED_DOC_EXT"])
            if saved_doc:
                player.document_path = saved_doc

        log_audit("PLAYER_PROFILE_UPDATED", "Player", str(player.id), f"Updated profile: {player.full_name}")
        db.session.commit()
        flash("Player profile updated successfully!", "success")
        return redirect(url_for("player_dashboard"))

    return render_template("player/edit_profile.html", player=player, districts=districts, clubs=clubs, categories=categories)


# ==============================================================================
# 7. COMPLETE ADMINISTRATIVE DASHBOARD ROUTES (/admin)
# ==============================================================================

@app.route("/admin", methods=["GET", "POST"])
@admin_required
def admin_dashboard():
    if request.method == "POST":
        action = request.form.get("action", "")

        # 1. Approve / Reject Player & Issue Permanent JKTTA ID
        if action == "approve_player":
            player = db.session.get(Player, request.form.get("player_id", type=int))
            if player:
                player.approval_status = "APPROVED"
                player.approved_at = utcnow()
                jktta_id = ensure_player_jktta_id(player)
                db.session.add(
                    Notification(
                        user_id=player.user_id,
                        title="Player Registration Approved!",
                        message=f"Congratulations! Your permanent JKTTA Player ID is {jktta_id}.",
                        link_url="/player/dashboard",
                    )
                )
                log_audit("PLAYER_APPROVED", "Player", str(player.id), f"Approved {player.full_name} -> {jktta_id}")
                db.session.commit()
                flash(f"Player approved! Assigned Permanent JKTTA ID: {jktta_id}", "success")

        elif action == "reject_player":
            player = db.session.get(Player, request.form.get("player_id", type=int))
            if player:
                player.approval_status = "REJECTED"
                player.review_remarks = request.form.get("remarks", "Application rejected by Secretariat.")
                log_audit("PLAYER_REJECTED", "Player", str(player.id), f"Rejected {player.full_name}")
                db.session.commit()
                flash("Player application rejected.", "warning")

        # 2. Add District
        elif action == "add_district":
            name = request.form.get("name", "").strip()
            code = request.form.get("code", "").strip().upper()
            division = request.form.get("division", "JAMMU").strip()
            if name and code:
                if not District.query.filter(or_(District.name == name, District.code == code)).first():
                    d = District(
                        name=name,
                        code=code,
                        division=division,
                        secretary_name=request.form.get("secretary_name", "").strip(),
                        contact_phone=request.form.get("contact_phone", "").strip(),
                    )
                    db.session.add(d)
                    log_audit("DISTRICT_ADDED", "District", code, f"Added district {name} ({code})")
                    db.session.commit()
                    flash(f"District {name} ({code}) added.", "success")
                else:
                    flash("District name or code already exists.", "warning")

        # 3. Add Club
        elif action == "add_club":
            name = request.form.get("name", "").strip()
            district_id = request.form.get("district_id", type=int)
            if name and district_id:
                c = Club(
                    name=name,
                    district_id=district_id,
                    head_coach=request.form.get("head_coach", "").strip(),
                    contact_phone=request.form.get("contact_phone", "").strip(),
                )
                db.session.add(c)
                log_audit("CLUB_ADDED", "Club", name, f"Added club {name}")
                db.session.commit()
                flash(f"Club '{name}' added successfully.", "success")

        # 4. Create Tournament + Initial Category
        elif action == "create_tournament":
            title = request.form.get("title", "").strip()
            venue = request.form.get("venue", "").strip()
            level = request.form.get("level", "STATE").strip()
            start_str = request.form.get("start_date", "")
            end_str = request.form.get("end_date", "")
            deadline_str = request.form.get("registration_deadline", "")
            cat_name = request.form.get("category_name", "Open Singles").strip()
            cat_gender = request.form.get("category_gender", "OPEN").strip()

            if title and venue and start_str and end_str and deadline_str:
                slug = f"{slugify_text(title)}-{secrets.token_hex(2)}"
                t = Tournament(
                    title=title,
                    slug=slug,
                    level=level,
                    venue=venue,
                    start_date=datetime.strptime(start_str, "%Y-%m-%d").date(),
                    end_date=datetime.strptime(end_str, "%Y-%m-%d").date(),
                    registration_deadline=datetime.strptime(deadline_str, "%Y-%m-%d").date(),
                    status="UPCOMING",
                    is_published=True,
                )
                db.session.add(t)
                db.session.flush()
                db.session.add(TournamentCategory(tournament_id=t.id, name=cat_name, gender=cat_gender))
                log_audit("TOURNAMENT_CREATED", "Tournament", str(t.id), f"Created tournament {title}")
                db.session.commit()
                flash(f"Tournament '{title}' created with category '{cat_name}'.", "success")

        # 5. Generate Knockout Fixtures
        elif action == "generate_fixtures":
            cat = db.session.get(TournamentCategory, request.form.get("category_id", type=int))
            if cat:
                try:
                    count = generate_knockout_fixtures(cat)
                    log_audit("FIXTURES_GENERATED", "TournamentCategory", str(cat.id), f"Generated {count} knockout matches")
                    db.session.commit()
                    flash(f"Generated {count} knockout matches for {cat.name}.", "success")
                except ValueError as exc:
                    flash(str(exc), "warning")

        # 6. Verify Match Result & Update Elo + Points
        elif action == "verify_match":
            match = db.session.get(Match, request.form.get("match_id", type=int))
            winner_id = request.form.get("winner_id", type=int)
            score_text = request.form.get("score_summary", "").strip()
            if match and winner_id in {match.player1_id, match.player2_id}:
                match.winner_id = winner_id
                match.status = "COMPLETED"
                match.is_verified = True
                if score_text:
                    MatchScore.query.filter_by(match_id=match.id).delete()
                    for idx, set_pair in enumerate(score_text.split(","), start=1):
                        parts = set_pair.strip().split("-")
                        if len(parts) == 2 and parts[0].isdigit() and parts[1].isdigit():
                            db.session.add(
                                MatchScore(
                                    match_id=match.id,
                                    set_number=idx,
                                    player1_points=int(parts[0]),
                                    player2_points=int(parts[1]),
                                )
                            )
                apply_verified_match_ratings(match)
                log_audit("MATCH_VERIFIED", "Match", str(match.id), f"Winner ID {winner_id} | Score: {score_text}")
                db.session.commit()
                flash("Match result verified and Elo / Association Points updated!", "success")

        # 7. Publish News Article
        elif action == "publish_news":
            title = request.form.get("title", "").strip()
            category_name = request.form.get("category_name", "Announcement").strip()
            summary = request.form.get("summary", "").strip()
            content = request.form.get("content", "").strip()
            if title and summary and content:
                img_file = request.files.get("featured_image")
                img_path = save_uploaded_file(img_file, "albums", app.config["ALLOWED_IMAGE_EXT"]) if img_file and img_file.filename else None
                article = NewsArticle(
                    title=title,
                    slug=f"{slugify_text(title)}-{secrets.token_hex(2)}",
                    category_name=category_name,
                    summary=summary,
                    content=content,
                    featured_image=img_path,
                    is_published=True,
                )
                db.session.add(article)
                log_audit("NEWS_PUBLISHED", "NewsArticle", title, f"Published news: {title}")
                db.session.commit()
                flash("News article published!", "success")

        # 8. Create Gallery Album & Upload Photos
        elif action == "upload_gallery":
            title = request.form.get("title", "").strip()
            category = request.form.get("category", "TOURNAMENT").strip()
            video_url = request.form.get("video_url", "").strip()
            files = request.files.getlist("photos")
            if title:
                album = GalleryAlbum(
                    title=title,
                    slug=f"{slugify_text(title)}-{secrets.token_hex(2)}",
                    category=category,
                    video_url=video_url,
                    is_published=True,
                )
                db.session.add(album)
                db.session.flush()
                saved_any = False
                for f in files:
                    p = save_uploaded_file(f, "albums", app.config["ALLOWED_IMAGE_EXT"])
                    if p:
                        if not album.cover_image:
                            album.cover_image = p
                        db.session.add(GalleryImage(album_id=album.id, image_path=p, caption=title))
                        saved_any = True
                if saved_any or video_url:
                    log_audit("GALLERY_ALBUM_CREATED", "GalleryAlbum", str(album.id), f"Created album {title}")
                    db.session.commit()
                    flash("Gallery album published!", "success")
                else:
                    db.session.rollback()
                    flash("Please upload at least one valid image (JPG/PNG/WEBP).", "danger")

        # 9. Update Association Secretariat & About Settings
        elif action == "update_settings":
            for field_key in (
                "contact_email",
                "contact_phone",
                "office_address",
                "president_name",
                "president_message",
                "secretary_name",
                "secretary_message",
                "about_history",
                "vision",
                "mission",
            ):
                val = request.form.get(field_key)
                if val is not None:
                    set_setting(field_key, val.strip())
            log_audit("SETTINGS_UPDATED", "SiteSetting", "global", "Updated association secretariat & about settings")
            db.session.commit()
            flash("Association settings updated.", "success")

        return redirect(url_for("admin_dashboard"))

    context = {
        "stats": {
            "total_players": Player.query.count(),
            "approved_players": Player.query.filter_by(approval_status="APPROVED").count(),
            "pending_players": Player.query.filter_by(approval_status="PENDING").count(),
            "districts": District.query.count(),
            "clubs": Club.query.count(),
            "tournaments": Tournament.query.count(),
        },
        "pending_players": Player.query.filter_by(approval_status="PENDING").order_by(Player.created_at.asc()).all(),
        "all_players": Player.query.order_by(Player.created_at.desc()).limit(30).all(),
        "districts": District.query.order_by(District.name.asc()).all(),
        "clubs": Club.query.order_by(Club.name.asc()).all(),
        "tournaments": Tournament.query.order_by(Tournament.start_date.desc()).all(),
        "open_matches": Match.query.filter(Match.status != "WALKOVER").order_by(Match.id.desc()).limit(25).all(),
        "audit_logs": AuditLog.query.order_by(AuditLog.created_at.desc()).limit(15).all(),
    }
    return render_template("admin/dashboard.html", **context)


# ==============================================================================
# 8. ERROR HANDLERS & CLI COMMANDS
# ==============================================================================

@app.errorhandler(403)
def err_403(e):
    return render_template("errors/403.html"), 403


@app.errorhandler(404)
def err_404(e):
    return render_template("errors/404.html"), 404


@app.errorhandler(500)
def err_500(e):
    db.session.rollback()
    return render_template("errors/500.html"), 500


@app.cli.command("init-db")
def init_db_cli():
    db.create_all()
    print("Database tables initialized successfully.")


@app.cli.command("create-admin")
def create_admin_cli():
    db.create_all()
    email = input("Super Admin Email: ").strip().lower()
    full_name = input("Full Name: ").strip()
    password = input("Password (min 8 chars): ").strip()
    if len(password) < 8:
        print("Error: Password must be at least 8 characters.")
        return
    if User.query.filter_by(email=email).first():
        print("Error: Email already registered.")
        return
    user = User(email=email, full_name=full_name, role="SUPER_ADMIN", is_active=True, is_email_verified=True)
    user.set_password(password)
    db.session.add(user)
    db.session.commit()
    print(f"Super Admin created: {email}")

def seed_official_ut_championship():
    slug = "49th-utt-jk-table-tennis-ut-championship-2026"
    if Tournament.query.filter_by(slug=slug).first():
        return

    t = Tournament(
        title="49th UTT J&K Table Tennis UT Championship - 2026",
        slug=slug,
        level="STATE",
        venue="Indoor Sports Complex, Table Tennis Hall, Moulana Azad Stadium, Jammu (J&K) - 180001",
        description=(
            "Conducted by Table Tennis Association Jammu & Kashmir (TTAJK) with Jammu & Kashmir Sports Council "
            "under the aegis of Table Tennis Federation of India (TTFI). Equipment: Stiga Tables & Butterfly 3 Star 40+ balls."
        ),
        entry_requirements=(
            "1. Only registered players of TTAJK for 2026 may participate.\n"
            "2. Age Cut-offs: U-19 (born on/after 01.01.2008), U-17 Junior (01.01.2010), "
            "U-15 Sub-Junior (01.01.2012), U-13 Cadet (01.01.2014), U-11 Hopes (01.01.2016).\n"
            "3. Original DOB certificate (Municipal Corporation / School) required for verification.\n"
            "4. Entries close on 7th December 2026 at 5:00 PM. Draw on 11th Dec 2026 at 6:00 AM at Venue."
        ),
        start_date=date(2026, 12, 9),
        end_date=date(2026, 12, 12),
        registration_deadline=date(2026, 12, 7),
        importance_multiplier=1.5,
        k_factor=32,
        status="UPCOMING",
        is_published=True,
    )
    db.session.add(t)
    db.session.flush()

    events = [
        ("Youth Boys Singles - U-11", "MALE", None, 11, 100),
        ("Youth Girls Singles - U-11", "FEMALE", None, 11, 100),
        ("Youth Boys Singles - U-13", "MALE", None, 13, 200),
        ("Youth Girls Singles - U-13", "FEMALE", None, 13, 200),
        ("Youth Boys Singles - U-15", "MALE", None, 15, 200),
        ("Youth Girls Singles - U-15", "FEMALE", None, 15, 200),
        ("Youth Boys Singles - U-17", "MALE", None, 17, 200),
        ("Youth Girls Singles - U-17", "FEMALE", None, 17, 200),
        ("Youth Boys Singles - U-19", "MALE", None, 19, 200),
        ("Youth Girls Singles - U-19", "FEMALE", None, 19, 200),
        ("Men Singles", "MALE", None, None, 300),
        ("Women's Singles", "FEMALE", None, None, 300),
        ("Veteran Men", "MALE", 40, None, 300),
    ]
    for cat_name, gender, min_a, max_a, fee in events:
        db.session.add(
            TournamentCategory(
                tournament_id=t.id,
                name=cat_name,
                gender=gender,
                min_age=min_a,
                max_age=max_a,
                entry_fee=fee,
            )
        )
    db.session.commit()


with app.app_context():
    db.create_all()
    seed_official_ut_championship()

if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)
