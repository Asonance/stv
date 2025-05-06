import pandas as pd
import math
from collections import defaultdict

def load_ballots(file_path):
    """Load ballots from Excel file"""
    df = pd.read_excel(file_path, header=None, skiprows=1)
    ballots = []
    for _, row in df.iterrows():
        # Remove NaN values and convert to list
        ballot = [c for c in row if pd.notna(c)]
        ballots.append(ballot)
    return ballots

def calculate_quota(ballots, seats):
    """Calculate the Droop quota"""
    return math.floor(len(ballots) / (seats + 1)) + 1

def count_first_preferences(ballots, active_candidates):
    """Count first preferences for active candidates"""
    counts = defaultdict(int)
    for ballot in ballots:
        for candidate in ballot:
            if candidate in active_candidates:
                counts[candidate] += 1
                break
    return counts

def stv_count(ballots, seats=5):
    """Main STV counting function"""
    elected = []
    eliminated = []
    all_candidates = list(set(c for ballot in ballots for c in ballot))
    active_candidates = all_candidates.copy()
    quota = calculate_quota(ballots, seats)
    
    print(f"Total ballots: {len(ballots)}")
    print(f"Quota for election: {quota}")
    print(f"Candidates: {', '.join(all_candidates)}\n")
    
    round_num = 1
    
    while len(elected) < seats and active_candidates:
        print(f"\n--- Round {round_num} ---")
        print(f"Active candidates: {', '.join(active_candidates)}")
        
        # Count current votes
        counts = count_first_preferences(ballots, active_candidates)
        sorted_counts = sorted(counts.items(), key=lambda x: (-x[1], x[0]))
        
        print("\nCurrent counts:")
        for candidate, votes in sorted_counts:
            print(f"{candidate}: {votes}")
        
        # Check for candidates reaching quota
        for candidate, votes in sorted_counts:
            if votes >= quota and candidate not in elected:
                elected.append(candidate)
                active_candidates.remove(candidate)
                print(f"\n{candidate} ELECTED with {votes} votes")
                
                # Redistribute surplus (simplified - just subtract quota)
                surplus = votes - quota
                if surplus > 0:
                    print(f"  Surplus of {surplus} votes to redistribute")
                    # In a full implementation, we would transfer these votes
                
                if len(elected) >= seats:
                    break
        
        if len(elected) >= seats:
            break
            
        # If no one elected, eliminate lowest candidate
        if not any(votes >= quota for _, votes in sorted_counts):
            # Find candidate(s) with lowest votes
            min_votes = min(counts.values())
            losers = [c for c, v in counts.items() if v == min_votes]
            
            # If tie, eliminate all with min votes (simplified)
            if len(losers) > 1:
                print(f"\nTIE: Multiple candidates with {min_votes} votes")
            
            for loser in losers:
                if loser not in eliminated:
                    eliminated.append(loser)
                    active_candidates.remove(loser)
                    print(f"\n{loser} ELIMINATED with {min_votes} votes")
                    
                    # In a full implementation, we would transfer these votes
        
        round_num += 1
    
    # If seats remain and candidates are left, elect remaining candidates
    remaining_seats = seats - len(elected)
    if remaining_seats > 0 and active_candidates:
        # Sort remaining by votes and elect top remaining
        counts = count_first_preferences(ballots, active_candidates)
        sorted_remaining = sorted(counts.items(), key=lambda x: (-x[1], x[0]))
        for candidate, _ in sorted_remaining[:remaining_seats]:
            elected.append(candidate)
            print(f"\n{candidate} ELECTED as remaining candidate")
    
    return elected

# Main execution
if __name__ == "__main__":
    # Load the ballots from Excel file
    file_path = "Copy of March 2024 Parliamentary Election (Raw Data).xlsx"
    ballots = load_ballots(file_path)
    
    # Run STV count (assuming 5 seats)
    seats = 5
    winners = stv_count(ballots, seats)
    
    print("\n=== FINAL RESULTS ===")
    print(f"Elected candidates (top {seats}):")
    for i, winner in enumerate(winners, 1):
        print(f"{i}. {winner}")