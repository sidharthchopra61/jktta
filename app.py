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
app.config["SECRET_KEY"] = os.getenv("SECRET_KEY", "jktta-local-dev-secret-change-in-production")

# Automatic runtime SQLite location (keeps project root strictly 4 items)
default_sqlite_dir = Path(tempfile.gettempdir()) / "jktta_runtime_db"
default_sqlite_dir.mkdir(parents=True, exist_ok=True)
default_db_uri = f"sqlite:///{(default_sqlite_dir / 'jktta.db').as_posix()}"

app.config["SQLALCHEMY_DATABASE_URI"] = os.getenv("DATABASE_URL", default_db_uri)
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["MAX_CONTENT_LENGTH"] = 5 * 1024 * 1024  # 5 MB max upload limit
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
    return re.sub(r"[\s_-]+", "-", cleaned)


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
    # Permanent JKTTA ID — identical across all tournaments for this player
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
    status = db.Column(db.String(30), default="PENDING", nullable=False, index=True)
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
# 3. JKTTA ID, TOURNAMENT BACK NUMBER & EXCEL EXPORT SERVICES
# ==============================================================================

def ensure_player_jktta_id(player: Player) -> str:
    """
    Assigns a permanent JKTTA ID once and keeps it identical across all tournaments.
    """
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
    """
    Generates a unique Back Number per tournament (e.g., 101, 102, 103...).
    Every player inside the same tournament receives a unique back number,
    and the same player can receive a different back number in another tournament.
    """
    existing_backs = (
        db.session.query(TournamentRegistration.back_number)
        .filter_by(tournament_id=tournament_id)
        .all()
    )
    used_numbers = set()
    for (bn,) in existing_backs:
        try:
            used_numbers.add(int(bn))
        except (ValueError, TypeError):
            pass

    candidate = 101
    while candidate in used_numbers:
        candidate += 1
    return str(candidate)


def build_tournament_excel_file(tournament_id: int) -> Path:
    """
    Generates/updates an official .xlsx roster file for the tournament inside gallery/documents/.
    """
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

    column_widths = [8, 12, 22, 26, 12, 15, 24, 18, 22, 15, 20]
    for col_idx, width in enumerate(column_widths, start=1):
        col_letter = ws.cell(row=3, column=col_idx).column_letter
        ws.column_dimensions[col_letter].width = width

    output_path = GALLERY_DIR / "documents" / f"tournament_{tournament.id}_players.xlsx"
    wb.save(str(output_path))
    return output_path


def register_player_for_tournament(player: Player, category: TournamentCategory) -> tuple[TournamentRegistration, Path]:
    existing = TournamentRegistration.query.filter_by(
        category_id=category.id,
        player_id=player.id,
    ).first()
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
        status="PENDING",
    )
    db.session.add(registration)
    db.session.commit()

    excel_path = build_tournament_excel_file(category.tournament_id)
    return registration, excel_path


# ==============================================================================
# 4. AUTHENTICATION, RBAC & TEMPLATE CONTEXT
# ==============================================================================

@login_manager.user_loader
def load_user(user_id):
    return db.session.get(User, int(user_id))


def role_required(*allowed_roles):
    def decorator(fn):
        @wraps(fn)
        @login_required
        def wrapper(*args, **kwargs):
            if current_user.role == "SUPER_ADMIN" or current_user.role in allowed_roles:
                return fn(*args, **kwargs)
            abort(403)
        return wrapper
    return decorator


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


@app.context_processor
def inject_globals():
    return {
        "jktta": {
            "ORG_NAME": get_setting("org_name", os.getenv("JKTTA_ORG_NAME", "Jammu & Kashmir Table Tennis Association")),
            "SHORT_NAME": get_setting("short_name", os.getenv("JKTTA_SHORT_NAME", "JKTTA")),
            "CONTACT_EMAIL": get_setting("contact_email", os.getenv("JKTTA_CONTACT_EMAIL", "info@jktta.example.org")),
            "CONTACT_PHONE": get_setting("contact_phone", os.getenv("JKTTA_CONTACT_PHONE", "+91-00000-00000")),
            "OFFICE_ADDRESS": get_setting(
                "office_address",
                os.getenv("JKTTA_OFFICE_ADDRESS", "Indoor Sports Complex, Jammu / Srinagar, J&K (Official Address Pending Verification)"),
            ),
            "PRESIDENT_NAME": get_setting("president_name", "To Be Updated by Secretariat"),
            "SECRETARY_NAME": get_setting("secretary_name", "To Be Updated by Secretariat"),
            "PRESIDENT_MESSAGE": get_setting(
                "president_message",
                "Official message from the President of JKTTA will be published here upon verification.",
            ),
            "SECRETARY_MESSAGE": get_setting(
                "secretary_message",
                "Official message from the General Secretary of JKTTA will be published here upon verification.",
            ),
            "ABOUT_HISTORY": get_setting(
                "about_history",
                "The Jammu & Kashmir Table Tennis Association coordinates district units, affiliated academies, state ranking tournaments, and national team selections across Jammu & Kashmir.",
            ),
            "VISION": get_setting(
                "vision",
                "To nurture disciplined, nationally competitive table tennis athletes from every district of Jammu & Kashmir through transparent governance and modern sports science.",
            ),
            "MISSION": get_setting(
                "mission",
                "To conduct regular district and UT championships, maintain auditable Elo and Association Points rankings, certify referees/coaches, and broaden grassroots access.",
            ),
            "MAP_EMBED_URL": get_setting("map_embed_url", os.getenv("JKTTA_MAP_EMBED_URL", "")),
        },
        "current_year": datetime.now().year,
    }


# ==============================================================================
# 5. MEDIA, PUBLIC PAGES (PHASE 1 + PHASE 2) & AUTH ROUTES
# ==============================================================================

@app.route("/gallery/<path:filename>")
def serve_gallery(filename):
    if filename.startswith("documents/"):
        if not current_user.is_authenticated:
            abort(403)
    return send_from_directory(app.config["GALLERY_FOLDER"], filename)


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


@app.route("/login", methods=["GET", "POST"])
def login():
    if current_user.is_authenticated:
        return redirect(url_for("home"))

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
            default_dest = url_for("player_dashboard") if user.player_profile else url_for("home")
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

# ==============================================================================
# PHASE 3: PLAYER REGISTRATION, EMAIL VERIFICATION, PASSWORD RESET & PLAYER PORTAL
# ==============================================================================

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

        if photo_file and photo_file.filename and not photo_path:
            flash("Invalid photo format. Allowed: PNG, JPG, JPEG, WEBP (max 5 MB).", "danger")
            return render_template("auth/register.html", districts=districts, clubs=clubs, categories=categories)

        if doc_file and doc_file.filename and not doc_path:
            flash("Invalid supporting document format. Allowed: PDF, PNG, JPG, JPEG (max 5 MB).", "danger")
            return render_template("auth/register.html", districts=districts, clubs=clubs, categories=categories)

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
                message="Your JKTTA player profile has been submitted and is awaiting District/Association verification.",
                link_url="/player/dashboard",
            )
        )
        log_audit("PLAYER_REGISTERED", "Player", str(player.id), f"New player application: {full_name} ({email})")
        db.session.commit()

        verify_link = url_for("verify_email", token=verify_token, _external=True)
        print(f"[JKTTA EMAIL VERIFICATION] {email} -> {verify_link}")
        flash(
            "Registration submitted! Your application status is PENDING admin approval. "
            "Check your terminal/inbox for your email verification link.",
            "success",
        )
        login_user(user)
        return redirect(url_for("player_dashboard"))

    return render_template("auth/register.html", districts=districts, clubs=clubs, categories=categories)


@app.route("/verify-email/<token>")
def verify_email(token):
    user = User.query.filter_by(verification_token=token).first()
    if not user:
        flash("Invalid or expired email verification link.", "danger")
        return redirect(url_for("home"))
    user.is_email_verified = True
    user.verification_token = None
    db.session.commit()
    flash("Email address verified successfully!", "success")
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
            reset_link = url_for("reset_password", token=token, _external=True)
            print(f"[JKTTA PASSWORD RESET] {email} -> {reset_link}")
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
        flash("Password updated successfully. Please log in.", "success")
        return redirect(url_for("login"))

    return render_template("auth/reset_password.html", token=token)


@app.route("/player/dashboard")
@login_required
def player_dashboard():
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
    results = (
        TournamentResult.query.filter_by(player_id=player.id)
        .order_by(TournamentResult.published_at.desc())
        .all()
    )
    achievements = Achievement.query.filter_by(player_id=player.id).order_by(Achievement.year.desc()).all()
    notifications = (
        Notification.query.filter_by(user_id=current_user.id)
        .order_by(Notification.created_at.desc())
        .limit(10)
        .all()
    )

    # Calculate profile completion percentage
    fields = [
        player.full_name,
        player.date_of_birth,
        player.gender,
        player.mobile,
        player.address,
        player.district_id,
        player.playing_category,
        player.photo_path,
        player.document_path,
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
        player.playing_hand = request.form.get("playing_hand", player.playing_hand).strip()
        player.playing_style = request.form.get("playing_style", player.playing_style).strip()
        player.guardian_name = request.form.get("guardian_name", "").strip()
        player.guardian_mobile = request.form.get("guardian_mobile", "").strip()

        photo_file = request.files.get("photo")
        if photo_file and photo_file.filename:
            saved_photo = save_uploaded_file(photo_file, "players", app.config["ALLOWED_IMAGE_EXT"])
            if saved_photo:
                player.photo_path = saved_photo
            else:
                flash("Invalid photo format. Allowed: PNG, JPG, JPEG, WEBP.", "danger")
                return render_template("player/edit_profile.html", player=player, districts=districts, clubs=clubs, categories=categories)

        doc_file = request.files.get("document")
        if doc_file and doc_file.filename:
            saved_doc = save_uploaded_file(doc_file, "documents", app.config["ALLOWED_DOC_EXT"])
            if saved_doc:
                player.document_path = saved_doc

        log_audit("PLAYER_PROFILE_UPDATED", "Player", str(player.id), f"Updated profile for {player.full_name}")
        db.session.commit()
        flash("Player profile updated successfully!", "success")
        return redirect(url_for("player_dashboard"))

    return render_template("player/edit_profile.html", player=player, districts=districts, clubs=clubs, categories=categories)



# ==============================================================================
# 6. ERROR HANDLERS & CLI COMMANDS
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


with app.app_context():
    db.create_all()

if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)